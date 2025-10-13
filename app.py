from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory, session
import os
import pandas as pd
from werkzeug.utils import secure_filename
from datetime import datetime, time
from openpyxl import Workbook, load_workbook
import plotly.express as px
import plotly.io as pio
import uuid

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
        sheet_name = request.form["sheet_name"]
        chart_type = request.form["chart_type"]
        x_col = request.form["x_column"]
        y_col = request.form["y_column"]
        operation = request.form["operation"]
        font = request.form["font"]
        color = request.form["color"]
        width = int(request.form.get("width", 800))
        height = int(request.form.get("height", 600))

        # ✅ Step 1: Map operation names to human-readable text
        operation_map = {
            "none": "",
            "sum": "Sum of",
            "avg": "Average of",
            "count": "Count of",
            "distinct_count": "Distinct Count of"
        }
        operation_text = operation_map.get(operation, "").strip()

        # ✅ Step 2: Build a smart, descriptive chart title
        if operation_text:
            chart_title = f"{sheet_name} — {operation_text} {y_col} by {x_col} ({chart_type.title()} Chart)"
        else:
            chart_title = f"{sheet_name} — {y_col} by {x_col} ({chart_type.title()} Chart)"

        # ✅ Aggregation
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

        # ✅ Keep y_col consistent for chart rendering
        y_col = df_result.columns[1]

        # ✅ Create chart
        custom_colors = [
            "#E57373", "#64B5F6", "#4DB6AC", "#FFD54F", "#81C784", "#FFB74D", "#9575CD"
        ]

        if chart_type == "bar":
            fig = px.bar(
                df_result,
                x=x_col,
                y=y_col,
                color=x_col,
                text=df_result[y_col].apply(lambda v: f"₹{v:,.0f}"),
                title=chart_title,
                color_discrete_sequence=custom_colors
            )
        else:
            fig = px.line(
                df_result,
                x=x_col,
                y=y_col,
                title=f"{sheet_name} - {chart_type.title()} Chart"
            )

        # ✅ Style chart
        fig.update_traces(
            texttemplate='%{text}',
            textposition='outside',
            hovertemplate='X: %{x}<br>Y: ₹%{y:,.0f}<extra></extra>'
        )
        fig.update_layout(
            width=width,
            height=height,
            uniformtext_minsize=12,
            uniformtext_mode='show',
            font=dict(family=font, size=14),
            title_font_color=color,
            xaxis_title_font_color=color,
            yaxis_title_font_color=color,
        )

        # ✅ Save chart as unique file (not in session)
        static_dir = os.path.join(BASE_DIR, "static")
        os.makedirs(static_dir, exist_ok=True)

        # Remove previous file if exists
        old_file = session.get("chart_file")
        if old_file:
            try:
                old_path = os.path.join(static_dir, old_file)
                if os.path.exists(old_path):
                    os.remove(old_path)
            except Exception:
                pass

        # Save new file
        new_filename = f"chart_{uuid.uuid4().hex}.html"
        new_path = os.path.join(static_dir, new_filename)
        fig.write_html(new_path, include_plotlyjs="cdn", full_html=True, config={'responsive': True})
        session["chart_file"] = new_filename

        # ✅ Log chart configuration
        log_df = pd.DataFrame([{
            "Sheet": sheet_name,
            "Chart Type": chart_type,
            "X Axis": x_col,
            "Y Axis": y_col,
            "Operation": operation,
            "Font": font,
            "Color": color
        }])
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            log_df.to_excel(writer, sheet_name="chart_config_log", index=False)

        flash("Chart configuration saved successfully.")
        return redirect(url_for("chart_view"))

    return render_template("data_model_2.html", columns=columns)


@app.route("/chart_view")
def chart_view():
    chart_file = session.get("chart_file", None)
    if chart_file:
        static_path = os.path.join(BASE_DIR, "static", chart_file)
        # wait until file fully written (max 2s)
        for _ in range(20):
            if os.path.exists(static_path) and os.path.getsize(static_path) > 1000:
                break
            time.sleep(0.1)
    return render_template("chart_view.html", chart_file=chart_file)


@app.route('/dashboard')
def dashboard():
    return render_template('dashboard.html', active_page='dashboard')

if __name__ == '__main__':
    app.run(debug=True, host="0.0.0.0")



