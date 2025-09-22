from flask import Flask, request, jsonify, render_template, send_file
from flask_cors import CORS
import pandas as pd
import io
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
latest_results_df = None
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def find_column(df, keywords):
    for col in df.columns:
        for keyword in keywords:
            if keyword.lower() in col.lower():
                return col
    # Fallback to column positions if no match found
    if 'sales' in keywords or 'sale' in keywords or 'qty' in keywords:
        if len(df.columns) > 1:
            return df.columns[1]
    if 'stock' in keywords or 'closing' in keywords or 'balance' in keywords:
        if len(df.columns) > 2:
            return df.columns[2]
    raise ValueError(f"No column found for {keywords}. Columns found: {list(df.columns)}")

def safe_float_convert(value, default=0.0):
    try:
        if pd.isna(value) or str(value).strip() == "":
            return default
        str_val = str(value).strip()
        if any(char.isalpha() for char in str_val):
            return default
        return float(str_val)
    except Exception:
        return default

def round_up_to_step(qty, step=5):
    if qty <= 0:
        return 0
    return int(((qty + step - 1) // step) * step)

def calculate_carton_order(order_qty, carton_size):
    if order_qty <= 0:
        return 0
    if order_qty <= carton_size:
        return carton_size
    full_cartons = int(order_qty // carton_size)
    remainder = order_qty % carton_size
    if remainder > 0:
        full_cartons += 1
    return full_cartons * carton_size

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    global latest_results_df
    try:
        if 'sales_file' not in request.files or 'stock_file' not in request.files:
            return jsonify({'error': 'Both sales and stock files are required'}), 400
        sales_file = request.files['sales_file']
        stock_file = request.files['stock_file']
        if sales_file.filename == '' or stock_file.filename == '':
            return jsonify({'error': 'No files selected'}), 400
        if not (allowed_file(sales_file.filename) and allowed_file(stock_file.filename)):
            return jsonify({'error': 'Only Excel files allowed (.xlsx, .xls)'}), 400
        sales_days = int(request.form.get('sales_days', 2))
        forecast_days = int(request.form.get('forecast_days', 2))
        selected_shop = request.form.get('selected_shop', 'Shop 01')
        sales_df = pd.read_excel(sales_file)
        stock_df = pd.read_excel(stock_file)
        results = process_inventory_data(sales_df, stock_df, sales_days, forecast_days, selected_shop)
        latest_results_df = pd.DataFrame(results)
        return jsonify({
            'success': True,
            'data': results,
            'selected_shop': selected_shop,
            'message': f'Processed {len(results)} items for {selected_shop} successfully'
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'Processing error: {str(e)}'}), 500

def process_inventory_data(sales_df, stock_df, sales_days, forecast_days, selected_shop):
    sales_df = sales_df.dropna(how='all').reset_index(drop=True)
    item_col = sales_df.columns[0]
    sales_col = find_column(sales_df, ['sale', 'sales', 'qty', 'sold', 'outward'])
    stock_col = find_column(sales_df, ['stock', 'closing', 'balance', 'current'])
    sales_df[sales_col] = sales_df[sales_col].apply(safe_float_convert)
    sales_df[stock_col] = sales_df[stock_col].apply(safe_float_convert)
    sales_df['daily_avg'] = sales_df[sales_col] / max(sales_days, 1)
    sales_df['forecast_qty'] = sales_df['daily_avg'] * forecast_days
    sales_df['order_qty'] = (sales_df['forecast_qty'] - sales_df[stock_col]).clip(lower=0)

    stock_columns = list(stock_df.columns)
    item_col_stock = stock_columns[0]
    carton_col = stock_columns[-1] if len(stock_columns) > 2 else None
    location_columns = [col for col in stock_columns[1:] if col != carton_col]

    results = []
    for _, row in sales_df.iterrows():
        item_name = str(row[item_col]).strip()
        if not item_name or item_name == 'nan':
            continue
        current_sale = float(row[sales_col])
        current_stock = float(row[stock_col])
        forecast_qty = float(row['forecast_qty'])
        order_qty = float(row['order_qty'])
        stock_match = stock_df[stock_df[item_col_stock].astype(str).str.strip() == item_name]
        if stock_match.empty:
            results.append(create_no_stock_result(item_name, current_sale, current_stock, forecast_qty, order_qty, selected_shop))
        else:
            results.append(process_item_with_stock_for_shop(stock_match.iloc[0], item_name, current_sale, current_stock, forecast_qty, order_qty, location_columns, carton_col, selected_shop))
    return results

def create_no_stock_result(item_name, current_sale, current_stock, forecast_qty, order_qty, selected_shop):
    return {
        'Item Name': item_name,
        'Sales': current_sale,
        'Stock': current_stock,
        'Command': int(order_qty) if order_qty > 0 else 0,
        'Order From Location': 'Not Available',
        'Available Qty': '0'
    }

def get_carton_size(stock_row, carton_col):
    carton_size = 1
    if carton_col and carton_col in stock_row.index:
        try:
            carton_size = int(float(stock_row[carton_col]))
            if carton_size <= 0: carton_size = 1
        except Exception:
            carton_size = 1
    return carton_size

def get_warehouse_stock(stock_row, location_columns):
    for col_name in location_columns:
        if 'warehouse' in str(col_name).lower() or 'wh' in str(col_name).lower():
            return safe_float_convert(stock_row[col_name], 0)
    return 0

def get_other_shop_stocks(stock_row, location_columns, selected_shop):
    other_shop_stocks = {}
    for col in location_columns:
        col_lower = str(col).lower()
        if ('warehouse' not in col_lower and col != selected_shop and 'shop' in col_lower):
            qty = safe_float_convert(stock_row[col], 0)
            if qty > 0:
                other_shop_stocks[col] = qty
    return other_shop_stocks

def get_max_shop_stock(shop_stocks):
    if shop_stocks:
        max_shop = max(shop_stocks.items(), key=lambda x: x[1])
        return max_shop[0], max_shop[1]
    return None, 0

def process_item_with_stock_for_shop(stock_row, item_name, current_sale, current_stock, forecast_qty, order_qty, location_columns, carton_col, selected_shop):
    carton_size = get_carton_size(stock_row, carton_col)
    warehouse_qty = get_warehouse_stock(stock_row, location_columns)
    other_shop_stocks = get_other_shop_stocks(stock_row, location_columns, selected_shop)
    max_shop_name, max_shop_qty = get_max_shop_stock(other_shop_stocks)

    if warehouse_qty >= order_qty:
        command_qty = int(round_up_to_step(order_qty, carton_size))
        command_source = "Warehouse"
        command_avail = str(int(warehouse_qty))
    elif max_shop_name and max_shop_qty >= order_qty:
        command_qty = int(round_up_to_step(order_qty, 5))
        command_source = max_shop_name
        command_avail = str(int(max_shop_qty))
    else:
        command_qty = int(round_up_to_step(order_qty, 5))
        command_source = "Insufficient Stock"
        if warehouse_qty > 0 and max_shop_qty > 0:
            command_avail = f"WH: {int(warehouse_qty)}, {max_shop_name}: {int(max_shop_qty)}"
        elif warehouse_qty > 0:
            command_avail = f"WH: {int(warehouse_qty)}"
        elif max_shop_qty > 0:
            command_avail = f"{max_shop_name}: {int(max_shop_qty)}"
        else:
            command_avail = "0"

    sources = []
    availabilities = []
    if warehouse_qty > 0:
        sources.append("Warehouse")
        availabilities.append(str(int(warehouse_qty)))
    if max_shop_name and max_shop_qty > 0:
        sources.append(max_shop_name)
        availabilities.append(str(int(max_shop_qty)))

    return {
        'Item Name': item_name,
        'Sales': current_sale,
        'Stock': current_stock,
        'Command': command_qty,
        'Order From Location': ", ".join(sources) if sources else command_source,
        'Available Qty': ", ".join(availabilities) if availabilities else command_avail
    }

@app.route('/export', methods=['POST'])
def export_filtered():
    data = request.get_json(force=True)
    df = pd.DataFrame(data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Filtered Orders')
    output.seek(0)
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        download_name='filtered_orders.xlsx',
        as_attachment=True
    )

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
