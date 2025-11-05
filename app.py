from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory, session, jsonify
import os
import pandas as pd
from werkzeug.utils import secure_filename
from datetime import datetime
import time
from openpyxl import Workbook, load_workbook
import plotly.express as px
import plotly.io as pio
import uuid
import numbers
import numpy as np

# Configuration
BASE_DIR = os.path.dirname(__file__)
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploaded_files')
EXCEL_FILE = os.path.join(BASE_DIR, 'file_log.xlsx')
SHEET_NAME = 'file header' # exactly as requested
# Set to None to accept any file type. To restrict, set to a set of extensions like {'pdf','png','jpg'}
ALLOWED_EXTENSIONS = None

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.secret_key = 'change-me-to-a-secure-random-string'

# ensure upload folder exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Retry helper (Windows file locks)
def _with_retry(func, retries=5, delay=0.3):
    for i in range(retries):
        try:
            return func()
        except PermissionError:
            if i == retries - 1:
                raise
            time.sleep(delay)

            
def allowed_file(filename):
    if ALLOWED_EXTENSIONS is None:
        return True
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def ensure_excel_sheet():
    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
        if "file header" not in wb.sheetnames:
            ws1 = wb.create_sheet("file header")
            ws1.append(['S.No', 'Path', 'file type', 'file name', 'file upload date', 'file upload time'])
        if "file details" not in wb.sheetnames:
            ws2 = wb.create_sheet("file details")
            ws2.append(['S.No', 'Path', 'file type', 'Sheet name', 'file name', 'file upload date', 'file upload time'])
    else:
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "file header"
        ws1.append(['S.No', 'Path', 'file type', 'file name', 'file upload date', 'file upload time'])
        ws2 = wb.create_sheet("file details")
        ws2.append(['S.No', 'Path', 'file type', 'Sheet name', 'file name', 'file upload date', 'file upload time'])
    return wb


@app.route('/')
def home():
    return render_template('index.html', active_page='home')

# @app.route('/upload')
# def upload():
#     return render_template('upload.html', active_page='upload')

@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        uploaded_file = request.files.get('file')
        if not uploaded_file or uploaded_file.filename == '':
            flash("Please select a file")
            return redirect(url_for('upload'))

        filename = secure_filename(uploaded_file.filename)
        timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
        saved_filename = f"{timestamp}_{filename}"
        save_path = os.path.join(app.config['UPLOAD_FOLDER'], saved_filename)
        uploaded_file.save(save_path)

        # Open or create Excel file
        wb = ensure_excel_sheet()

        # Common values
        rel_path = os.path.relpath(save_path, BASE_DIR)
        file_type = uploaded_file.mimetype
        today = datetime.now().strftime("%Y-%m-%d")
        now_time = datetime.now().strftime("%H:%M:%S")

        # 1️⃣ Log into file header
        ws1 = wb["file header"]
        s_no_header = ws1.max_row  # since first row is header
        ws1.append([s_no_header, rel_path, file_type, filename, today, now_time])

        # 2️⃣ Log into file details (one row per sheet in uploaded Excel)
        try:
            uploaded_wb = load_workbook(save_path, read_only=True)
            ws2 = wb["file details"]
            for sheet in uploaded_wb.sheetnames:
                s_no_detail = ws2.max_row  # since first row is header
                ws2.append([s_no_detail, rel_path, file_type, sheet, filename, today, now_time])
            uploaded_wb.close()
        except Exception as e:
            flash(f"Could not read sheets from uploaded file: {e}")

        # Save and close
        wb.save(EXCEL_FILE)
        wb.close()

        flash("File uploaded successfully")
        return redirect(url_for('configuration'))

    return render_template('upload.html', active_page='upload')


# @app.route('/configuration')
# def configuration():
#     return render_template('configuration.html', active_page='configuration')

@app.route('/configuration', methods=['GET', 'POST'])
def configuration():
    import os
    from openpyxl import Workbook, load_workbook
    from flask import flash, redirect, url_for, request, render_template

    # 1️⃣ Collect all sheet names from uploaded Excel files
    sheet_names = set()  # use set to avoid duplicates
    if os.path.exists(EXCEL_FILE):
        wb_log = load_workbook(EXCEL_FILE, read_only=True)
        if "file details" in wb_log.sheetnames:
            ws_details = wb_log["file details"]
            for row in ws_details.iter_rows(min_row=2, values_only=True):
                sheet_in_file = row[3]  # column 3 = Sheet Name
                if sheet_in_file:
                    sheet_names.add(sheet_in_file)
        wb_log.close()
    sheet_names = sorted(list(sheet_names))  # convert set to sorted list

    if request.method == 'POST':
        # 2️⃣ Get all table names and selected sheets from the form
        table_names = request.form.getlist("tableName[]")
        selected_sheets = request.form.getlist("sheetName[]")

        if os.path.exists(EXCEL_FILE):
            wb = load_workbook(EXCEL_FILE)
        else:
            wb = Workbook()

        # 3️⃣ Create or reset "configuration log" sheet
        config_sheet_name = "configuration log"
        if config_sheet_name in wb.sheetnames:
            ws_config = wb[config_sheet_name]
            wb.remove(ws_config)
        ws_config = wb.create_sheet(config_sheet_name)

        # 4️⃣ Write header
        ws_config.append(["S.No", "Path", "File Name", "Sheet Name"])

        # 5️⃣ Log selected sheets
        if "file details" in wb.sheetnames:
            ws_details = wb["file details"]
            file_rows = list(ws_details.iter_rows(min_row=2, values_only=True))
            for idx, (table_name, sheet_name) in enumerate(zip(table_names, selected_sheets), start=1):
                # Find the corresponding file info from file details
                for row in file_rows:
                    path, file_name, sheet_in_file = row[1], row[4], row[3]
                    if sheet_in_file == sheet_name:
                        ws_config.append([idx, path, file_name, sheet_name])
                        break

        # 6️⃣ Save workbook
        wb.save(EXCEL_FILE)
        wb.close()

        flash("Configuration saved successfully!")
        return redirect(url_for('report_configuration'))

    # 7️⃣ Render template with sheet names for dropdown
    return render_template('configuration.html', active_page='configuration', sheet_names=sheet_names)




@app.route('/report-configuration', methods=['GET', 'POST'])
def report_configuration():
    # Read configuration log (sheets selected in previous step)
    config_rows = []
    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
        if "configuration log" in wb.sheetnames:
            ws = wb["configuration log"]
            for r in ws.iter_rows(min_row=2, values_only=True):
                if any(r):
                    config_rows.append({
                        "sno": r[0],
                        "path": r[1],
                        "file_name": r[2],
                        "sheet_name": r[3]
                    })
        wb.close()

    if request.method == 'POST':
        header_map = request.form.to_dict()   # { "sheetname": "row_number", ... }

        wb = load_workbook(EXCEL_FILE)
        # Create or get file_header_columns sheet
        if "file_header_columns" in wb.sheetnames:
            ws_cols = wb["file_header_columns"]
        else:
            ws_cols = wb.create_sheet("file_header_columns")
            ws_cols.append(["S.No", "Path", "File Name", "Sheet Name", "Column Name", "SheetName_ColumnName"])

        s_no = ws_cols.max_row  # continue numbering

        for cfg in config_rows:
            sheet_name = cfg["sheet_name"]
            header_row = int(header_map.get(sheet_name, 1))  # default row=1
            file_path = os.path.join(BASE_DIR, cfg["path"])
            file_name = cfg["file_name"]

            # open uploaded Excel file
            if os.path.exists(file_path):
                wb_file = load_workbook(file_path, read_only=True)
                if sheet_name in wb_file.sheetnames:
                    ws_file = wb_file[sheet_name]
                    headers = [cell.value for cell in ws_file[header_row] if cell.value]
                    for col in headers:
                        s_no += 1
                        ws_cols.append([
                            s_no,
                            cfg["path"],
                            file_name,
                            sheet_name,
                            col,
                            f"{sheet_name}_{col}"
                        ])
                wb_file.close()

        wb.save(EXCEL_FILE)
        wb.close()
        flash("Report configuration saved. Columns extracted successfully!")
        return redirect(url_for('data_model'))

    return render_template('report_configuration.html', active_page='report_configuration', config_rows=config_rows)


# Data Model page (join builder) - accepts both /data-model and /data_model
@app.route('/data-model', methods=['GET', 'POST'])
@app.route('/data_model', methods=['GET', 'POST'])
def data_model():
    # Load available columns from file_header_columns sheet
    available_columns = {}      # { "SheetName": [ {"col": column_name, "option": option_value}, ... ] }
    option_map = {}             # { "SheetName_ColumnName": (sheet, column) }
    sheet_to_path = {}          # { "SheetName": relative_path }

    if os.path.exists(EXCEL_FILE):
        try:
            wb = load_workbook(EXCEL_FILE, read_only=True)
            if "file_header_columns" in wb.sheetnames:
                ws = wb["file_header_columns"]
                for r in ws.iter_rows(min_row=2, values_only=True):
                    # expected: [S.No, Path, File Name, Sheet Name, Column Name, SheetName_ColumnName]
                    path = r[1]; file_name = r[2]; sheet = r[3]; col = r[4]; opt = r[5]
                    if sheet and opt:
                        available_columns.setdefault(sheet, []).append({"col": col, "option": opt})
                        option_map[opt] = (sheet, col)
                        # track the path for that sheet (first occurrence)
                        if sheet not in sheet_to_path and path:
                            sheet_to_path[sheet] = path
            wb.close()
        except Exception as e:
            flash(f"Error reading file_header_columns: {e}")
            return render_template('data_model.html', active_page='data_model',
                                   available_columns=available_columns)

    if request.method == 'POST':
        left_opts = request.form.getlist("left_col[]")
        right_opts = request.form.getlist("right_col[]")
        join_types = request.form.getlist("join_type[]")

        # Basic validation
        if not (left_opts and right_opts and join_types) or not (len(left_opts) == len(right_opts) == len(join_types)):
            flash("Invalid form submission: ensure each row has left, right and join type.")
            return redirect(url_for('data_model'))

        # Prepare dataframes cache per (path, sheet) so we don't reload repeatedly
        df_cache = {}  # key = (path, sheet) -> DataFrame

        def load_sheet_df(sheet_name):
            """Load the dataframe for that sheet using sheet_to_path mapping."""
            if sheet_name not in sheet_to_path:
                return None, f"No file path found for sheet '{sheet_name}' in file_header_columns."
            path_rel = sheet_to_path[sheet_name]
            path_abs = os.path.join(BASE_DIR, path_rel)
            if not os.path.exists(path_abs):
                return None, f"File not found: {path_abs}"
            key = (path_abs, sheet_name)
            if key in df_cache:
                return df_cache[key], None
            try:
                df = pd.read_excel(path_abs, sheet_name=sheet_name, engine="openpyxl")
                df_cache[key] = df
                return df, None
            except Exception as e:
                return None, f"Failed to read {sheet_name} from {path_abs}: {e}"

        # perform sequential joins (chain must be connected)
        result_df = None
        included_sheets = set()
        join_map = {"Inner Join": "inner", "Left Join": "left", "Right Join": "right", "Outer Join": "outer"}

        join_log_rows = []  # to record what was joined for the data_model sheet

        try:
            for l_opt, r_opt, jt in zip(left_opts, right_opts, join_types):
                if l_opt not in option_map or r_opt not in option_map:
                    flash(f"Column mapping not found for {l_opt} or {r_opt}")
                    return redirect(url_for('data_model'))

                l_sheet, l_col = option_map[l_opt]
                r_sheet, r_col = option_map[r_opt]

                df_left, err = load_sheet_df(l_sheet)
                if err:
                    flash(err); return redirect(url_for('data_model'))
                df_right, err = load_sheet_df(r_sheet)
                if err:
                    flash(err); return redirect(url_for('data_model'))

                how = join_map.get(jt, "inner")

                if result_df is None:
                    # first join
                    if l_col not in df_left.columns:
                        flash(f"Column '{l_col}' not in sheet '{l_sheet}'"); return redirect(url_for('data_model'))
                    if r_col not in df_right.columns:
                        flash(f"Column '{r_col}' not in sheet '{r_sheet}'"); return redirect(url_for('data_model'))
                    result_df = pd.merge(df_left, df_right, left_on=l_col, right_on=r_col, how=how, suffixes=('_L','_R'))
                    included_sheets.update([l_sheet, r_sheet])
                else:
                    # subsequent join: determine which side (left/right) is already in result_df
                    # we check presence of columns named exactly l_col or r_col in result_df
                    if l_col in result_df.columns:
                        # result_df has left column -> merge with df_right using r_col
                        if r_col not in df_right.columns:
                            flash(f"Column '{r_col}' not found in {r_sheet}"); return redirect(url_for('data_model'))
                        result_df = pd.merge(result_df, df_right, left_on=l_col, right_on=r_col, how=how, suffixes=('','_R2'))
                        included_sheets.add(r_sheet)
                    elif r_col in result_df.columns:
                        # result_df has right column -> merge with df_left
                        if l_col not in df_left.columns:
                            flash(f"Column '{l_col}' not found in {l_sheet}"); return redirect(url_for('data_model'))
                        result_df = pd.merge(result_df, df_left, left_on=r_col, right_on=l_col, how=how, suffixes=('','_L2'))
                        included_sheets.add(l_sheet)
                    else:
                        # not connected -> abort with clear message
                        flash(f"Join chain is not connected for {l_opt} <> {r_opt}. Make sure joins form a connected chain.")
                        return redirect(url_for('data_model'))

                join_log_rows.append((l_opt, r_opt, jt))

        except Exception as e:
            flash(f"Unexpected error during join: {e}")
            return redirect(url_for('data_model'))

        # If nothing was joined
        if result_df is None:
            flash("No joins performed.")
            return redirect(url_for('data_model'))
        
        # Clean NaN and NaT values before saving
        result_df = result_df.fillna("")  # replaces NaN, NaT, None with empty string


        # Save joined result to EXCEL_FILE as 'joined_output' (replace if exists)
        try:
            if os.path.exists(EXCEL_FILE):
                with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    result_df.to_excel(writer, sheet_name='joined_output', index=False)
            else:
                with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
                    result_df.to_excel(writer, sheet_name='joined_output', index=False)
        except Exception as e:
            flash(f"Failed to write joined_output sheet: {e}")
            return redirect(url_for('data_model'))

        # Log the join definitions into 'data_model' sheet
        try:
            wb_log = load_workbook(EXCEL_FILE)
            if "data_model" not in wb_log.sheetnames:
                ws_dm = wb_log.create_sheet("data_model")
                ws_dm.append(["S.No", "Sheet1_Column", "Sheet2_Column", "Join_Type"])
            else:
                ws_dm = wb_log["data_model"]
            start_idx = ws_dm.max_row
            for i, (lopt, ropt, jt) in enumerate(join_log_rows, start=start_idx):
                ws_dm.append([i, lopt, ropt, jt])
            wb_log.save(EXCEL_FILE)
            wb_log.close()
        except Exception as e:
            flash(f"Failed to log data_model details: {e}")
            return redirect(url_for('data_model'))

        # Save the joined DataFrame temporarily as JSON in session or just redirect to preview route
        # We'll just redirect to preview route; preview route will read 'joined_output' sheet.
        flash("Join completed and saved to 'joined_output'. Redirecting to preview...")
        return redirect(url_for('data_model_2'))

    # GET: render builder UI
    return render_template('data_model.html', active_page='data_model', available_columns=available_columns)



# --- page 2: preview joined output ---
@app.route("/data_model_2", methods=["GET", "POST"])
def data_model_2():
    joined_sheet = "joined_output"

    if not os.path.exists(EXCEL_FILE):
        flash("Excel file not found. Please upload again.")
        return redirect(url_for("data_model"))

    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=joined_sheet)
    except Exception as e:
        flash(f"Error loading joined_output: {e}")
        return redirect(url_for("data_model"))

    columns = list(df.columns)

    if request.method == "POST":

        print("Selected columns raw:", request.form.get("selected_columns"))


        # -----------------------------
        # Step 1: Collect all chart configurations
        # -----------------------------
        charts_data = []
        index = 0
        while f"chart_title_{index}" in request.form:
            chart_title = request.form.get(f"chart_title_{index}")
            chart_type = request.form.get(f"chart_type_{index}")
            x_col = request.form.get(f"x_column_{index}")
            y_col = request.form.get(f"y_column_{index}")
            operation = request.form.get(f"operation_{index}")
            font = request.form.get(f"font_{index}")
            color = request.form.get(f"color_{index}")
            width = int(request.form.get(f"width_{index}", 800))
            height = int(request.form.get(f"height_{index}", 600))

            charts_data.append({
                "chart_title": chart_title,
                "chart_type": chart_type,
                "x_col": x_col,
                "y_col": y_col,
                "operation": operation,
                "font": font,
                "color": color,
                "width": width,
                "height": height
            })
            index += 1

            print("Chart config:", chart_type, font, color, width, height)

        # -----------------------------
        # Step 2: Generate each chart
        # -----------------------------
        generated_files = []
        custom_colors = [
            "#E57373", "#64B5F6", "#4DB6AC", "#FFD54F", "#81C784", "#FFB74D", "#9575CD"
        ]
        operation_map = {
            "none": "",
            "sum": "Sum of",
            "avg": "Average of",
            "count": "Count of",
            "distinct_count": "Distinct Count of"
        }

        
        for chart_conf in charts_data:
            chart_title = chart_conf["chart_title"]
            chart_type = chart_conf["chart_type"]
            x_col = chart_conf["x_col"]
            y_col = chart_conf["y_col"]
            operation = chart_conf["operation"]
            font = chart_conf["font"]
            color = chart_conf["color"]
            width = chart_conf["width"]
            height = chart_conf["height"]

            operation_text = operation_map.get(operation, "").strip()
            custom_title = chart_conf.get("chart_title", "").strip()

            if custom_title:
                chart_title = custom_title
            else:
                if chart_type == "card":
                    chart_title = f"Summary of {y_col}"
                elif chart_type == "pie":
                    chart_title = f"{y_col} Distribution by {x_col} ({chart_type.title()} Chart)"
                elif chart_type == "table":
                    chart_title = f"Data Table"
                elif chart_type == "matrix_table":
                    chart_title = f"Data Table - Matrix(Summable)"
                else:
                    if operation_text:
                        chart_title = f"{operation_text} {y_col} by {x_col} ({chart_type.title()} Chart)"
                    else:
                        chart_title = f"{y_col} by {x_col} ({chart_type.title()} Chart)"

            # ---------- Aggregation ----------
            if operation == "sum":
                df_result = df.groupby(x_col)[y_col].sum().reset_index()
            elif operation == "avg":
                df_result = df.groupby(x_col)[y_col].mean().reset_index()
            elif operation == "count":
                df_result = df.groupby(x_col)[y_col].count().reset_index()
                df_result.rename(columns={y_col: f"Count of {y_col}"}, inplace=True)
            elif operation == "distinct_count":
                df_result = df.groupby(x_col)[y_col].nunique().reset_index()
                df_result.rename(columns={y_col: f"Distinct Count of {y_col}"}, inplace=True)
            else:
                df_result = df[[x_col, y_col]]

            y_col_actual = df_result.columns[1]
            # Check if we should show currency
            use_currency = operation in ["sum", "avg"]

            # ---------- Create Chart ----------
            if chart_type == "bar":
                fig = px.bar(
                    df_result,
                    x=x_col,
                    y=y_col_actual,
                    color=x_col,
                    text=df_result[y_col_actual].apply(lambda v: f"₹{v:,.0f}" if use_currency else f"{v}"),
                    title=chart_title,
                    color_discrete_sequence=custom_colors
                )
                fig.update_traces(
                    texttemplate='%{text}',
                    textposition='outside',
                    textangle=90,
                    cliponaxis=False,
                    hovertemplate=f'X: %{{x}}<br>Y: {"₹%{y:,.0f}" if use_currency else "%{y}"}<extra></extra>'
                )
                # Dynamic y-axis with extra space
                max_value = df_result[y_col_actual].max()
                fig.update_yaxes(
                    range=[0, max_value*1.2], dtick=5000
                )
                chart_title = f"<b>{chart_title}</b>"
                fig.update_layout(
                    title=dict(
                        text=chart_title,
                        font=dict(size=25, color=color, family=font),
                        x=0.5,
                        xanchor='center'
                    ),
                    width=width,
                    height=height,
                    uniformtext_minsize=12,
                    uniformtext_mode='show',
                    font=dict(family=font, size=10),
                    title_font_color=color,
                    xaxis_title_font_color=color,
                    yaxis_title_font_color=color,
                    bargap=0.3,
                    bargroupgap=0.1,
                    margin=dict(l=50, r=50, t=80, b=80)
                )
            elif chart_type == "line":
                fig = px.line(
                    df_result,
                    x=x_col,
                    y=y_col_actual,
                    title=chart_title,
                    markers=True
                )
                fig.update_traces(
                    texttemplate=f'{"₹%{y:,.0f}" if use_currency else "%{y}"}',
                    textposition='top center',
                    hovertemplate=f'X: %{{x}}<br>Y: {"₹%{y:,.0f}" if use_currency else "%{y}"}',
                    mode='lines+markers+text'
                )
                chart_title = f"<b>{chart_title}</b>"
                fig.update_layout(
                    title=dict(
                        text=chart_title,
                        font=dict(size=25, color=color, family=font),
                        x=0.5,
                        xanchor='center'
                    ),
                    width=width,
                    height=height, 
                    font=dict(family=font, size=10), 
                    title_font_color=color,
                    xaxis_title_font_color=color,
                    yaxis_title_font_color=color,
                    uniformtext_minsize=12,
                    uniformtext_mode='show',
                    margin=dict(l=50, r=50, t=80, b=80)
                )

            elif chart_type == "pie":
                percentages = df_result[y_col_actual] / df_result[y_col_actual].sum() * 100
                text_positions = ['inside' if p >= 5 else 'outside' for p in percentages]
                pull_values = [0.05 if p < 5 else 0.05 for p in percentages]
                text_templates = [
                    f"%{{label}}: {'₹' if use_currency else ''}%{{value:,.0f}} ({p:.1f}%)" if p >= 5 else f"{'₹' if use_currency else ''}%{{value:,.0f}} ({p:.1f}%)"
                    for p in percentages
                ]
                fig = px.pie(
                    df_result,
                    names=x_col,
                    values=y_col_actual,
                    title=chart_title,
                    color_discrete_sequence=custom_colors
                )
                fig.update_traces(
                    textposition=text_positions,
                    texttemplate=text_templates,
                    insidetextorientation='radial',
                    textfont_size=10,
                    pull=pull_values,
                    marker_line_width=1,
                    hovertemplate=f"%{{label}}: {'₹' if use_currency else ''}%{{value:,.0f}} (%{{percent}})<extra></extra>"
                )
                chart_title = f"<b>{chart_title}</b>"
                fig.update_layout(
                    title=dict(
                        text=chart_title,
                        font=dict(size=25, color=color, family=font),
                        x=0.5,
                        xanchor='center'
                    ),
                    width=width, 
                    height=height, 
                    font=dict(family=font, size=10), 
                    title_font_color=color,
                    margin=dict(t=50,b=50,l=50,r=50), 
                    showlegend=False
                    )
            elif chart_type == "card":
                total_value = df_result[y_col_actual].sum() if operation in ["sum", "none"] else df_result[y_col_actual].mean()
                card_html = f"""
                <div style='
                    display:flex;
                    flex-direction:column;
                    align-items:center;
                    justify-content:center;
                    height:300px;
                    background:linear-gradient(135deg, {color}, #f8f9fa);
                    border-radius:16px;
                    font-family:{font};
                    box-shadow:0 4px 10px rgba(0,0,0,0.1);
                '>
                    <h2 style='color:#333; margin-bottom:10px;'>{chart_title}</h2>
                    <h1 style='font-size:3rem; color:{color}; margin:0;'>₹{total_value:,.0f}</h1>
                </div>
                """

            elif chart_type == "table":
                selected_cols_str = request.form.get(f"selected_columns_{charts_data.index(chart_conf)}", "")
                selected_cols = [col.strip() for col in selected_cols_str.split(",") if col.strip()]

                # Filter only columns that exist in df_result
                existing_cols = [col for col in selected_cols if col in df.columns]
                missing_cols = [col for col in selected_cols if col not in df.columns]

                if missing_cols:
                    print(f"Warning: These columns are missing from df_result: {missing_cols}")

                if not existing_cols:
                    card_html = "<p style='color:red;'>⚠️ No valid columns selected for table.</p>"
                else:
                    df_table = df[existing_cols]

                    # Add serial numbers (1, 2, 3, ...)
                    df_table.insert(0, "S.No", range(1, len(df_table) + 1))

                    # ---------- Format currency columns ----------
                    amount_keywords = ["amount", "total", "transaction", "price", "cost"]
                    currency_cols = [
                        col for col in df_table.columns 
                        if any(keyword.lower() in col.lower() for keyword in amount_keywords)
                    ]
                    for col in currency_cols:
                        df_table[col] = df_table[col].apply(
                            lambda x: f"₹{int(x):,}" if pd.notnull(x) else ""
                            )


                    # Wrap table in a container div and center it
                    table_html = f"""
                    <div style="overflow-x:auto; display:flex; justify-content:center; align-items:center; margin:auto;">
                        {df_table.to_html(
                            classes="table table-striped table-bordered table-sm text-center",
                            index=False,
                            escape=False
                        )}
                    </div>
                    """
                    table_html = f"""
                    <style>
                        table th, table td {{
                            text-align: center;
                        }}
                    </style>
                    {table_html}
                    """

                    card_html = f"""
                    <div style='
                        background:#fff;
                        border-radius:12px;
                        box-shadow:0 4px 8px rgba(0,0,0,0.1);
                        overflow:auto;
                        padding:15px;
                    '>
                    <h4 style='
                        font-family:{font}; 
                        color:{color}; 
                        font-size:25px; 
                        text-align:center;
                        font-weight:bold;
                        margin-bottom:22px;
                        '>{chart_title}</h4>
                        {table_html}
                    </div>
                    """
                    

            elif chart_type == "matrix_table":
                selected_cols_str = request.form.get(f"selected_columns_{charts_data.index(chart_conf)}", "")
                selected_cols = [col.strip() for col in selected_cols_str.split(",") if col.strip()]

                # Filter only columns that exist in df_result
                existing_cols = [col for col in selected_cols if col in df.columns]
                missing_cols = [col for col in selected_cols if col not in df.columns]

                if missing_cols:
                    print(f"Warning: These columns are missing from df_result: {missing_cols}")

                if not existing_cols:
                    card_html = "<p style='color:red;'>⚠️ No valid columns selected for table.</p>"
                else:

                    # Create df_table with only selected columns
                    df_table = df[existing_cols]

                    # Add serial numbers (1, 2, 3, ...)
                    df_table.insert(0, "S.No", range(1, len(df_table) + 1))

                    # Define keywords for summable columns
                    summable_keywords = ["amount", "total", "qty", "quantity", "price", "value", "cost"]

                    # Calculate totals
                    totals = {}
                    for col in df_table.columns:
                        col_lower = col.lower()
                        if col == "S.No":
                            totals[col] = "Total"
                        elif any(keyword in col_lower for keyword in summable_keywords):
                            if pd.api.types.is_numeric_dtype(df_table[col]):
                                totals[col] = df_table[col].sum()
                            else:
                                totals[col] = ""
                        else:
                            totals[col] = "" 

                    # Append total row
                    df_table.loc[len(df_table)] = totals


                    # ---------- Format currency columns ----------
                    amount_keywords = ["amount", "total", "price", "cost"]
                    currency_cols = [
                        col for col in df_table.columns 
                        if any(keyword.lower() in col.lower() for keyword in amount_keywords)
                    ]
                    for col in currency_cols:
                        df_table[col] = df_table[col].apply(
                            lambda x: f"₹{int(x):,}" if pd.notnull(x) else ""
                            )
                    

                    # Wrap table in a container div and center it
                    table_html = f"""
                    <div style="overflow-x:auto; display:flex; justify-content:center; align-items:center; margin:auto;">
                        {df_table.to_html(
                            classes="table table-striped table-bordered table-sm text-center",
                            index=False,
                            escape=False
                        )}
                    </div>
                    """
                    # Apply bold + background to last row (Total) using CSS inside <style>
                    table_html = f"""
                    <style>
                        table tbody tr:last-child {{
                            font-weight: bold;
                            background-color: #f0f0f0;
                        }}
                        table th, table td {{
                            text-align: center;
                        }}
                    </style>
                    {table_html}
                    """

                    card_html = f"""
                    <div style='
                        background:#fff;
                        border-radius:12px;
                        box-shadow:0 4px 8px rgba(0,0,0,0.1);
                        overflow:auto;
                        padding:15px;
                    '>
                    <h4 style='
                        font-family:{font}; 
                        color:{color}; 
                        font-size:25px; 
                        text-align:center;
                        font-weight:bold;
                        margin-top:14px;
                        margin-bottom:22px;
                        '>
                        {chart_title}</h4>
                        {table_html}
                    </div>
                    """

            else:
                continue  # skip unknown chart type



            # ---------- Save chart ----------
            static_dir = os.path.join(BASE_DIR, "static")
            os.makedirs(static_dir, exist_ok=True)
            new_filename = f"chart_{uuid.uuid4().hex}.html"
            new_path = os.path.join(static_dir, new_filename)

            if chart_type in ["card", "table", "matrix_table"]:
                full_html = f"""
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset="utf-8">
                    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
                    <style>
                        body {{
                            margin: 0;
                            padding: 10px;
                            background: #fafafa;
                        }}
                    </style>
                </head>
                <body>
                    <div class="chart-wrapper" id="chart-{uuid.uuid4().hex}" style="margin-bottom:20px;">
                        {card_html}
                    </div>
                </body>
                </html>
                """
                with open(new_path, "w", encoding="utf-8") as f:
                    f.write(full_html)
            else:
                fig.write_html(new_path, include_plotlyjs="cdn", full_html=True, config={'responsive': True})

                # --- Ensure DOCTYPE at top to prevent Quirks Mode ---
                with open(new_path, "r+", encoding="utf-8") as f:
                    content = f.read()
                    if not content.lstrip().startswith("<!DOCTYPE html>"):
                        f.seek(0)
                        f.write("<!DOCTYPE html>\n" + content)
                        f.truncate()

            generated_files.append(new_filename)

        # -----------------------------
        # Step 3: Store generated files in session
        # -----------------------------
        session["chart_files"] = generated_files
        session["charts_meta"] = charts_data
        print(f"{len(generated_files)} chart(s) generated successfully.")
        return redirect(url_for("chart_view"))

    return render_template("data_model_2.html", columns=columns)



@app.route("/chart_view", methods=["GET"])
def chart_view():
    joined_sheet = "joined_output"

    # load DF (same way you already do)
    if not os.path.exists(EXCEL_FILE):
        flash("Excel file not found.")
        return redirect(url_for("data_model"))

    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=joined_sheet)
    except Exception as e:
        flash(f"Error loading joined_output: {e}")
        return redirect(url_for("data_model"))

    
    

    # ----------- Clean NaN, NaT, None values -----------
    df_cleaned = df.copy()

    # Replace NaN/NaT/pd.NA with empty string
    df_cleaned = df_cleaned.replace({pd.NA: "", np.nan: "", "NaT": "", pd.NaT: ""})
    
    # Handle datetime columns cleanly
    for col in df_cleaned.select_dtypes(include=["datetime", "datetimetz"]).columns:
        df_cleaned[col] = df_cleaned[col].astype(str).replace("NaT", "")

    df_cleaned = df_cleaned.where(pd.notnull(df_cleaned), None)
    joined_data = df_cleaned.to_dict(orient="records")
    print(len(joined_data), "rows prepared for chart_view.")

    # charts_meta stored earlier in session
    charts_meta = session.get("charts_meta", [])  # list of chart config dicts
    print(f"Loaded {len(charts_meta)} chart metadata entries from session.")

    # chart_files still used by your current UI if needed
    chart_files = session.get("chart_files", [])
    
    # build a map of unique values per column (strings) for the filter UI
    column_values = {}
    for col in df.columns:
        # dropna then unique and convert to python types (string or number)
        vals = df[col].dropna().unique().tolist()
        # convert numpy types to Python native
        vals = [v.item() if hasattr(v, "item") else v for v in vals]
        # sort strings, numbers works too
        try:
            vals_sorted = sorted(vals, key=lambda x: (str(type(x)), x))
        except Exception:
            vals_sorted = vals
        column_values[col] = vals_sorted

    # ensure static files ready (your existing waiting logic)
    if chart_files:
        static_paths = [os.path.join(BASE_DIR, "static", chart_file) for chart_file in chart_files]
        for _ in range(20):
            if all(os.path.exists(static_path) and os.path.getsize(static_path) > 1000 for static_path in static_paths):
                break
            time.sleep(0.1)

    return render_template(
        "chart_view.html",
        chart_files=chart_files,
        charts_meta=charts_meta,
        joined_data=joined_data,
        column_values=column_values
    )




@app.route("/table_view")
def table_view():
    joined_sheet = "joined_output"

    if not os.path.exists(EXCEL_FILE):
        flash("Excel file not found. Please upload again.")
        return redirect(url_for("data_model"))

    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=joined_sheet)
    except Exception as e:
        flash(f"Error loading joined_output: {e}")
        return redirect(url_for("data_model"))
    
    # ----------- Replace NaN & NaT with empty string ----------
    df = df.replace({pd.NA: "", np.nan: "", "NaT": "", pd.NaT: ""})
    

    # ----------- Fix float columns ----------
    for col in df.columns:
        if df[col].dtype == 'float64':
            # If the column is float but should be int
            if df[col].dropna().apply(lambda x: float(x).is_integer() if pd.notnull(x) else True).all():
                df[col] = df[col].astype('Int64').astype(str).replace("<NA>", "")

    # ----------- Add Serial Number (S.No) ----------
    df.insert(0, "S.No", range(1, len(df)+1))

    # ----------- Format amount/currency columns ----------
    amount_keywords = ["amount", "total", "price", "cost", "transaction_amount"]
    currency_cols = [
        col for col in df.columns if any(keyword.lower() in col.lower() for keyword in amount_keywords)
    ]
    for col in currency_cols:
        def format_currency(x):
            try:
                # Try converting to float
                num = float(str(x).replace(",", "").replace("₹", "").strip())
                return f"₹{int(num):,}"
            except (ValueError, TypeError):
                # If it's empty, non-numeric, or invalid — return blank
                return ""
        df[col] = df[col].apply(format_currency)

    # ----------- Add totals row for amount columns ----------
    totals = {}
    for col in df.columns:
        if col in currency_cols:
            # Sum numeric values before formatting
            col_numeric_sum = pd.to_numeric(df[col].str.replace('₹','').str.replace(',',''), errors='coerce').sum()
            totals[col] = f"₹{int(col_numeric_sum):,}"
        elif col == "S.No":
            totals[col] = "Total"
        else:
            totals[col] = ""

    # Append total row
    df.loc[len(df)] = totals

    # --- Pagination ---
    page = int(request.args.get('page', 1))
    per_page = 20
    # Calculate total pages
    total_rows = len(df) - 1  # minus totals row
    total_pages = (total_rows + per_page - 1) // per_page
    # Convert to list of dicts for Jinja rendering
    data = df.to_dict(orient='records')
    start_idx = (page - 1) * per_page
    end_idx = start_idx + per_page
    page_data = data[start_idx:end_idx]
    
    

    # Convert df to HTML table
    table_html = df.to_html(
        classes="table table-striped table-bordered table-sm text-center",
        index=False,
        escape=False
    )

    # # Apply CSS for total row bold + background
    # table_html = f"""
    # <style>
    #     table tbody tr:last-child {{
    #         font-weight: bold;
    #         background-color: #f0f0f0;
    #     }}
    #     table th, table td {{
    #         text-align: center;
    #     }}
    # </style>
    # {table_html}
    # """


    # CSS for total row if shown
    # style_html = ""
    # if show_totals:
    #     style_html = """
    #     <style>
    #         table tbody tr:last-child {
    #             font-weight: bold;
    #             background-color: #f0f0f0;
    #         }
    #         table th, table td {
    #             text-align: center;
    #         }
    #     </style>
    #     """
    # else:
    #     style_html = """
    #     <style>
    #         table th, table td {
    #             text-align: center;
    #         }
    #     </style>
    #     """

    

    return render_template("table_view.html", 
                           data=page_data,
                           columns=df.columns, 
                           page=page, 
                           total_pages=total_pages)


@app.route('/dashboard')
def dashboard():
    return render_template('dashboard.html', active_page='dashboard')

if __name__ == '__main__':
    app.run(debug=True, host="0.0.0.0")



