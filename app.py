import csv
import io
import re
import secrets
from collections import defaultdict
from datetime import datetime

import pandas as pd
from flask import Flask, render_template, request, send_file, abort

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

MAX_GROUP_SIZE = 8

# Temporary storage for individual Excel downloads.
DOWNLOADS = {}

REQUIRED_COLUMNS = [
    "Student Name",
    "Chest Number",
    "School Name",
    "Category",
    "Sports Event",
]

COLUMN_ALIASES = {
    "studentname": "Student Name",
    "name": "Student Name",
    "nameofstudent": "Student Name",
    "participant": "Student Name",
    "chest": "Chest Number",
    "chestno": "Chest Number",
    "chestnum": "Chest Number",
    "chestnumber": "Chest Number",
    "chestnumberno": "Chest Number",
    "schoolname": "School Name",
    "school": "School Name",
    "category": "Category",
    "agecategory": "Category",
    "sportsevent": "Sports Event",
    "event": "Sports Event",
    "eventname": "Sports Event",
    "item": "Sports Event",
}


def normalize_column_name(name):
    return re.sub(r"[^a-z0-9]", "", str(name).strip().lower())


def get_normalized_aliases():
    return {
        normalize_column_name(alias): target
        for alias, target in COLUMN_ALIASES.items()
    }


def find_header_row(rows):
    aliases = get_normalized_aliases()

    for row_index, row in enumerate(rows[:50]):
        found = set()

        for value in row:
            target = aliases.get(normalize_column_name(value))
            if target:
                found.add(target)

        if all(column in found for column in REQUIRED_COLUMNS):
            return row_index

    return None


def standardize_columns(df):
    aliases = get_normalized_aliases()
    rename_map = {}

    for column in df.columns:
        target = aliases.get(normalize_column_name(column))

        if target:
            rename_map[column] = target

    df = df.rename(columns=rename_map)
    df = df.loc[:, ~df.columns.duplicated()].copy()

    missing = [
        column for column in REQUIRED_COLUMNS
        if column not in df.columns
    ]

    if missing:
        raise ValueError(
            "Missing required column(s): "
            + ", ".join(missing)
            + ". Please check your spreadsheet headings."
        )

    return df


def read_csv_with_header(content, filename):
    try:
        text = content.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = content.decode("latin-1")

    best_rows = None
    best_header_index = None
    best_score = -1

    for delimiter in [",", "\t", ";", "|"]:
        try:
            rows = list(
                csv.reader(io.StringIO(text), delimiter=delimiter)
            )
        except csv.Error:
            continue

        header_index = find_header_row(rows)

        if header_index is None:
            continue

        header_width = len(rows[header_index])

        if header_width > best_score:
            best_rows = rows
            best_header_index = header_index
            best_score = header_width

    if best_rows is None:
        raise ValueError(
            f"'{filename}' does not contain a recognizable header row "
            "with all five required columns. Check that the file includes "
            "Participant, Chest, School, Category, and Item headings."
        )

    headers = [
        str(value).strip()
        for value in best_rows[best_header_index]
    ]

    data_rows = best_rows[best_header_index + 1:]
    width = len(headers)
    cleaned_rows = []

    for row in data_rows:
        row = list(row[:width])

        if len(row) < width:
            row.extend([""] * (width - len(row)))

        cleaned_rows.append(row)

    return pd.DataFrame(cleaned_rows, columns=headers)


def read_excel_with_header(content, filename, extension):
    engine = "xlrd" if extension == "xls" else "openpyxl"

    raw_df = pd.read_excel(
        io.BytesIO(content),
        sheet_name=0,
        header=None,
        dtype=str,
        keep_default_na=False,
        engine=engine,
    )

    rows = raw_df.fillna("").values.tolist()
    header_index = find_header_row(rows)

    if header_index is None:
        raise ValueError(
            f"'{filename}' does not contain a recognizable header row "
            "with all five required columns."
        )

    headers = [
        str(value).strip()
        for value in raw_df.iloc[header_index].tolist()
    ]

    df = raw_df.iloc[header_index + 1:].copy()
    df.columns = headers

    return df


def read_uploaded_file(file):
    filename = file.filename or ""
    extension = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""

    if extension not in {"xlsx", "xls", "csv"}:
        raise ValueError(
            f"'{filename}' is not supported. Upload an XLSX, XLS, or CSV file."
        )

    content = file.read()

    if not content:
        raise ValueError(f"'{filename}' is empty.")

    try:
        if extension == "csv":
            df = read_csv_with_header(content, filename)
        else:
            df = read_excel_with_header(content, filename, extension)

    except ValueError:
        raise
    except Exception as exc:
        raise ValueError(
            f"Could not read '{filename}'. Check that the file is valid. "
            f"Details: {exc}"
        ) from exc

    if df.empty:
        raise ValueError(f"'{filename}' contains no usable student records.")

    df = standardize_columns(df)

    for column in REQUIRED_COLUMNS:
        df[column] = (
            df[column].fillna("").astype(str).str.strip()
        )

    df = df[
        df[
            ["Student Name", "Chest Number", "School Name", "Sports Event"]
        ].apply(lambda row: any(value for value in row), axis=1)
    ].copy()

    if df.empty:
        raise ValueError(f"'{filename}' contains no usable student records.")

    if (df["Student Name"] == "").any():
        raise ValueError(
            f"'{filename}' has a student record with a missing Student Name."
        )

    if (df["School Name"] == "").any():
        raise ValueError(
            f"'{filename}' has a student record with a missing School Name."
        )

    if (df["Sports Event"] == "").any():
        raise ValueError(
            f"'{filename}' has a student record with a missing Sports Event."
        )

    categories = df.loc[
        df["Category"] != "", "Category"
    ].unique().tolist()

    if not categories:
        raise ValueError(
            f"'{filename}' has no category values in the Category column."
        )

    if len(categories) > 1:
        raise ValueError(
            f"'{filename}' contains multiple categories: "
            + ", ".join(categories)
            + ". Each input file must contain only one category."
        )

    df["Category"] = categories[0]

    return df


def create_groups(event_df):
    """
    Randomly group students:
    - Maximum 8 students per group.
    - No two students from the same school in one group.
    """
    students = event_df.to_dict("records")
    secrets.SystemRandom().shuffle(students)

    groups = []

    for student in students:
        school = student["School Name"]

        eligible_groups = [
            group
            for group in groups
            if len(group["students"]) < MAX_GROUP_SIZE
            and school not in group["schools"]
        ]

        if eligible_groups:
            smallest_size = min(
                len(group["students"])
                for group in eligible_groups
            )

            least_loaded = [
                group
                for group in eligible_groups
                if len(group["students"]) == smallest_size
            ]

            selected_group = secrets.choice(least_loaded)

        else:
            selected_group = {
                "students": [],
                "schools": set(),
            }

            groups.append(selected_group)

        selected_group["students"].append(student)
        selected_group["schools"].add(school)

    return groups


def safe_filename(name):
    name = str(name).strip()
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    name = name.rstrip(". ")

    if not name:
        name = "Uncategorized"

    if name.upper() in {
        "CON", "PRN", "AUX", "NUL",
        *(f"COM{i}" for i in range(1, 10)),
        *(f"LPT{i}" for i in range(1, 10)),
    }:
        name = "_" + name

    return name[:100]


def safe_sheet_name(name, used_names):
    name = re.sub(r"[\[\]:*?/\\]", "_", str(name)).strip()
    name = name[:31] or "Event"

    candidate = name
    counter = 2

    while candidate.casefold() in used_names:
        suffix = f"_{counter}"
        candidate = name[:31 - len(suffix)] + suffix
        counter += 1

    used_names.add(candidate.casefold())

    return candidate


def build_category_workbook(category, category_df):
    output = io.BytesIO()
    summary_rows = []
    used_sheet_names = {"summary"}

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        events = sorted(
            category_df["Sports Event"].unique(),
            key=lambda value: value.casefold(),
        )

        for event in events:
            event_df = category_df[
                category_df["Sports Event"] == event
            ].copy()

            groups = create_groups(event_df)
            output_rows = []

            for group_number, group in enumerate(groups, start=1):
                group_students = group["students"]

                if len(group_students) > MAX_GROUP_SIZE:
                    raise RuntimeError("A group exceeded the maximum size.")

                school_names = [
                    student["School Name"]
                    for student in group_students
                ]

                if len(school_names) != len(set(school_names)):
                    raise RuntimeError(
                        "A group contains two students from the same school."
                    )

                for student in group_students:
                    output_rows.append({
                        "Group Number": group_number,
                        "Student Name": student["Student Name"],
                        "Chest Number": student["Chest Number"],
                        "School Name": student["School Name"],
                        "Category": category,
                        "Sports Event": event,
                    })

                summary_rows.append({
                    "Category": category,
                    "Sports Event": event,
                    "Group Number": group_number,
                    "Student Count": len(group_students),
                    "Schools": ", ".join(school_names),
                })

            result_df = pd.DataFrame(
                output_rows,
                columns=[
                    "Group Number",
                    "Student Name",
                    "Chest Number",
                    "School Name",
                    "Category",
                    "Sports Event",
                ],
            )

            sheet_name = safe_sheet_name(event, used_sheet_names)

            result_df.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False,
            )

        summary_df = pd.DataFrame(
            summary_rows,
            columns=[
                "Category",
                "Sports Event",
                "Group Number",
                "Student Count",
                "Schools",
            ],
        )

        summary_df.to_excel(
            writer,
            sheet_name="Summary",
            index=False,
        )

        workbook = writer.book

        for worksheet in workbook.worksheets:
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions

            for column_cells in worksheet.columns:
                max_length = 0
                column_letter = column_cells[0].column_letter

                for cell in column_cells:
                    value = "" if cell.value is None else str(cell.value)
                    max_length = max(max_length, len(value))

                worksheet.column_dimensions[column_letter].width = min(
                    max(max_length + 2, 12), 45
                )

    output.seek(0)

    return output.getvalue(), summary_rows


@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "GET":
        return render_template("index.html")

    files = request.files.getlist("files")
    files = [file for file in files if file and file.filename]

    if not files:
        return render_template(
            "index.html",
            error="Please select at least one Excel or CSV file.",
        ), 400

    try:
        category_frames = defaultdict(list)
        uploaded_count = 0

        for file in files:
            df = read_uploaded_file(file)
            category = df["Category"].iloc[0]

            category_frames[category].append(df)
            uploaded_count += 1

        combined_categories = {
            category: pd.concat(frames, ignore_index=True)
            for category, frames in category_frames.items()
        }

        results = []
        used_filenames = set()

        for category, category_df in combined_categories.items():

            workbook_bytes, summary_rows = build_category_workbook(
                category,
                category_df,
            )

            base_name = safe_filename(category)
            filename = f"{base_name}.xlsx"
            counter = 2

            while filename.casefold() in used_filenames:
                filename = f"{base_name}_{counter}.xlsx"
                counter += 1

            used_filenames.add(filename.casefold())

            token = secrets.token_urlsafe(24)

            DOWNLOADS[token] = {
                "data": workbook_bytes,
                "filename": filename,
            }

            results.append({
                "category": category,
                "students": len(category_df),
                "events": category_df["Sports Event"].nunique(),
                "groups": len(summary_rows),
                "filename": filename,
                "download_token": token,
            })

        # Keep only the latest 50 generated downloads.
        while len(DOWNLOADS) > 50:
            oldest_token = next(iter(DOWNLOADS))
            del DOWNLOADS[oldest_token]

        total_students = sum(
            item["students"] for item in results
        )

        total_groups = sum(
            item["groups"] for item in results
        )

        return render_template(
            "results.html",
            results=results,
            total_students=total_students,
            total_groups=total_groups,
            category_count=len(results),
            uploaded_count=uploaded_count,
        )

    except ValueError as exc:
        return render_template(
            "index.html",
            error=str(exc),
        ), 400

    except Exception as exc:
        app.logger.exception("SportSort processing failed")

        return render_template(
            "index.html",
            error=f"Something went wrong while processing the files: {exc}",
        ), 500


@app.route("/download/<token>")
def download(token):

    item = DOWNLOADS.get(token)

    if item is None:
        abort(
            404,
            description="This download has expired. Please generate it again."
        )

    return send_file(
        io.BytesIO(item["data"]),
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        as_attachment=True,
        download_name=item["filename"],
    )


@app.errorhandler(413)
def file_too_large(_error):

    return render_template(
        "index.html",
        error="The uploaded files are too large. The total upload limit is 50 MB.",
    ), 413


if __name__ == "__main__":
    app.run(debug=True)