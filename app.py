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
latest_full_results = None
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def find_column(df, keywords):
    for col in df.columns:
        for keyword in keywords:
            if str(keyword).lower() in str(col).lower():
                return col
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
    global latest_results_df, latest_full_results
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

        # Read first 2500 rows, then keep only first 11 columns (safe for any file)
        sales_df = pd.read_excel(sales_file, nrows=2500)
        sales_df = sales_df.iloc[:, :11]
        stock_df = pd.read_excel(stock_file, nrows=2500)
        stock_df = stock_df.iloc[:, :11]

        results = process_inventory_data(sales_df, stock_df, sales_days, forecast_days, selected_shop)
        latest_results_df = pd.DataFrame(results)
        latest_full_results = results
        return jsonify({
            'success': True,
            'data': results,
            'selected_shop': selected_shop,
            'message': f'Processed {len(results)} items for {selected_shop} successfully (limited to first 2500 rows & max 11 columns)'
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

    all_items = set(sales_df[item_col].astype(str).str.strip()) | set(stock_df[item_col_stock].astype(str).str.strip())

    for item_name in all_items:
        item_name = str(item_name).strip()
        if not item_name or str(item_name).lower() == 'nan':
            continue

        sales_row = sales_df[sales_df[item_col].astype(str).str.strip() == item_name]
        stock_row = stock_df[stock_df[item_col_stock].astype(str).str.strip() == item_name]

        if sales_row.empty and not stock_row.empty:
            row = stock_row.iloc[0]
            current_sale = 0.0
            current_stock = safe_float_convert(row[location_columns[0]]) if location_columns else 0.0
            if current_stock < 10:
                forecast_qty = 0
                order_qty = 10 - current_stock
                results.append(process_item_with_stock_for_shop(row, item_name, current_sale, current_stock, forecast_qty, order_qty, location_columns, carton_col, selected_shop))
            else:
                results.append(create_no_stock_result(item_name, current_sale, current_stock, 0, 0, selected_shop))

        elif not sales_row.empty:
            row = sales_row.iloc[0]
            current_sale = float(row[sales_col])
            current_stock = float(row[stock_col])
            forecast_qty = float(row['forecast_qty'])
            order_qty = float(row['order_qty'])
            if current_sale == 0 and current_stock < 10:
                forecast_qty = 0
                order_qty = 10 - current_stock
            stock_match = stock_df[stock_df[item_col_stock].astype(str).str.strip() == item_name]
            if stock_match.empty:
                results.append(create_no_stock_result(item_name, current_sale, current_stock, forecast_qty, order_qty, selected_shop))
            else:
                results.append(process_item_with_stock_for_shop(stock_match.iloc[0], item_name, current_sale, current_stock, forecast_qty, order_qty, location_columns, carton_col, selected_shop))
        else:
            results.append(create_no_stock_result(item_name, 0, 0, 0, 0, selected_shop))

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
        if ('warehouse' not in col_lower and str(col) != str(selected_shop) and 'shop' in col_lower):
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
    step_size = 5
    warehouse_qty = get_warehouse_stock(stock_row, location_columns)
    other_shop_stocks = get_other_shop_stocks(stock_row, location_columns, selected_shop)
    max_shop_name, max_shop_qty = get_max_shop_stock(other_shop_stocks)

    if warehouse_qty >= order_qty:
        command_qty = int(round_up_to_step(order_qty, step_size))
        command_source = "Warehouse"
        command_avail = str(int(warehouse_qty))
    elif max_shop_name and max_shop_qty >= order_qty:
        command_qty = int(round_up_to_step(order_qty, step_size))
        command_source = max_shop_name
        command_avail = str(int(max_shop_qty))
    else:
        command_qty = int(round_up_to_step(order_qty, step_size))
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
        sources.append(str(max_shop_name))
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
    global latest_full_results
    filtered_data = request.get_json(force=True)
    filtered_df = pd.DataFrame(filtered_data)
    all_df = pd.DataFrame(latest_full_results if latest_full_results is not None else filtered_data)
    
    if not filtered_df.empty and 'Item Name' in filtered_df.columns:
        filtered_df = filtered_df.sort_values('Item Name')
    if not all_df.empty and 'Item Name' in all_df.columns:
        all_df = all_df.sort_values('Item Name')
    
    total_items = len(all_df)
    low_stock_count = (all_df['Stock'].apply(lambda x: float(str(x).replace(",", "")) if str(x).replace(",", "").replace(".", "").isdigit() else 0) < 10).sum()
    zero_stock_count = (all_df['Stock'].apply(lambda x: float(str(x).replace(",", "")) if str(x).replace(",", "").replace(".", "").isdigit() else 0) == 0).sum()
    order_needed_count = (all_df['Command'].apply(lambda x: float(str(x).replace(",", "")) if str(x).replace(",", "").replace(".", "").isdigit() else 0) > 0).sum()
    summary_df = pd.DataFrame([
        {"Metric": "Total Items", "Value": total_items},
        {"Metric": "Low Stock (<10)", "Value": int(low_stock_count)},
        {"Metric": "Zero Stock (0)", "Value": int(zero_stock_count)},
        {"Metric": "Order Needed (Command > 0)", "Value": int(order_needed_count)}
    ])

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        filtered_df.to_excel(writer, index=False, sheet_name='Filtered Orders')
        all_df.to_excel(writer, index=False, sheet_name='All Items')
        summary_df.to_excel(writer, index=False, sheet_name='Summary')

        workbook = writer.book

        header_format = workbook.add_format({
            'bold': True,
            'font_color': 'white',
            'font_size': 12,
            'bg_color': '#3E50B4',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        item_format = workbook.add_format({
            'font_name': 'Calibri',
            'font_size': 12,
            'border': 1,
            'align': 'left',
            'valign': 'vcenter'
        })
        number_format = workbook.add_format({
            'font_name': 'Calibri',
            'font_size': 12,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '0'
        })
        cell_format = workbook.add_format({
            'font_name': 'Calibri',
            'font_size': 12,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        summary_cell_format = workbook.add_format({
            'border': 1,
            'font_name': 'Calibri',
            'font_size': 12,
            'align': 'left'
        })

        for sheet_name, df in [
            ('Filtered Orders', filtered_df),
            ('All Items', all_df),
            ('Summary', summary_df)
        ]:
            worksheet = writer.sheets[sheet_name]
            (max_row, max_col) = df.shape
            if max_col > 0:
                worksheet.autofilter(0, 0, max_row, max_col - 1)
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    worksheet.set_column(col_num, col_num, 19)
                for row in range(1, max_row + 1):
                    for col in range(max_col):
                        if sheet_name in ['Filtered Orders', 'All Items'] and col == 0:
                            worksheet.write(row, col, df.iloc[row-1, col], item_format)
                        elif sheet_name in ['Filtered Orders', 'All Items'] and col in [1, 2, 3]:
                            try:
                                val = int(float(df.iloc[row-1, col]))
                            except:
                                val = df.iloc[row-1, col]
                            worksheet.write(row, col, val, number_format)
                        elif sheet_name == 'Summary':
                            worksheet.write(row, col, df.iloc[row-1, col], summary_cell_format)
                        else:
                            worksheet.write(row, col, df.iloc[row-1, col], cell_format)
    output.seek(0)
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        download_name='advanced_orders_export.xlsx',
        as_attachment=True
    )

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
