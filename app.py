from pydoc import doc
import streamlit as st
import zipfile
import io
import json
import os
from pathlib import Path
from docx import Document
from pbixray import PBIXRay

# Parse the uploaded .pbix file using PBIXRay
def parse_pbixray(uploaded_zip):
    # Use PBIXRay to parse the Power BI file
    file_bytes = uploaded_zip.getvalue()
    model = PBIXRay(io.BytesIO(uploaded_zip.getvalue()))
    
    tables = model.tables
    if hasattr(tables, "to_dict"):
        tables = tables.to_dict(orient="records")  # Convert to list of dicts
    metadata = model.metadata
    if hasattr(metadata, "to_dict"):
        metadata = metadata.to_dict(orient="records")  # Convert to list of dicts
    power_query = model.power_query
    if hasattr(power_query, "to_dict"):
        power_query = power_query.to_dict(orient="records")  # Convert to list of dicts
    m_parameters = model.m_parameters
    if hasattr(m_parameters, "to_dict"):
        m_parameters = m_parameters.to_dict(orient="records")  # Convert to list of dicts
    dax_tables = model.dax_tables
    if hasattr(dax_tables, "to_dict"):
        dax_tables = dax_tables.to_dict(orient="records")  # Convert to list of dicts
    measures = model.dax_measures
    if hasattr(measures, "to_dict"):
        measures = measures.to_dict(orient="records")  # Convert to list of dicts
    dax_columns = model.dax_columns
    if hasattr(dax_columns, "to_dict"):
        dax_columns = dax_columns.to_dict(orient="records")  # Convert to list of dicts
    aggregations = model.aggregations
    if hasattr(aggregations, "to_dict"):
        aggregations = aggregations.to_dict(orient="records")  # Convert to list of dicts
    schema = model.schema
    if hasattr(schema, "to_dict"):
        schema = schema.to_dict(orient="records")  # Convert to list of dicts
    relationships = model.relationships
    if hasattr(relationships, "to_dict"):
        relationships = relationships.to_dict(orient="records")  # Convert to list of dicts
    rls = model.rls
    if hasattr(rls, "to_dict"):
        rls = rls.to_dict(orient="records")  # Convert to list of dicts
    connections = model.connections
    if hasattr(connections, "to_dict"):
        connections = connections.to_dict(orient="records")  # Convert to list of dicts
    
    return {
        "tables": tables,
        "metadata": metadata,
        "power_query": power_query,
        "m_parameters": m_parameters,
        "dax_tables": dax_tables,
        "measures": measures,
        "dax_columns": dax_columns,
        "aggregations": aggregations,
        "schema": schema,
        "relationships": relationships,
        "rls": rls,
        "connections": connections
    }

# Extract visuals from the uploaded .pbix file
def extract_visuals_from_pbix(uploaded_zip):
    file_bytes = uploaded_zip.getvalue()
    # pull the JSON files from the PBIX archive and parse them to extract visual information
    visuals = []
    warnings = []

    try:
        # Open the PBIX file as a zip archive and iterate through its contents
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
            # Iterate through the files in the PBIX archive
            for name in z.namelist():
                # loop through the files in the PBIX archive and look for JSON files
                lower = name.lower()
                # lowcase the name of the file and check if it ends with .json, if not, continue to the next file
                if not lower.endswith(".json"):
                    continue
                try:
                    raw = z.read(name)
                    text = raw.decode("utf-8", errors="ignore")
                    payload = json.loads(text)
                except Exception:
                    continue

                def walk(obj):
                    if isinstance(obj, dict):
                        # A common report/visual shape in embedded PBIX JSON
                        if "visualType" in obj or "visualContainer" in obj or "name" in obj:
                            item = {}

                            if "name" in obj:
                                item["name"] = obj.get("name")
                            if "visualType" in obj:
                                item["visualType"] = obj.get("visualType")
                            if "x" in obj and "y" in obj and "width" in obj and "height" in obj:
                                item["layout"] = {
                                    "x": obj.get("x"),
                                    "y": obj.get("y"),
                                    "width": obj.get("width"),
                                    "height": obj.get("height"),
                                }
                            # pull filters from the same visual object
                            if "filters" in obj:
                                item["filters"] = obj.get("filters")

                            # pull columns from the same object
                            if "query" in obj:
                                query = obj.get("query", {})
                                if isinstance(query, dict):
                                    item["columns"] = query.get("columns", [])

                            #pull raw data transforms if present
                            if "dataTransforms" in obj:
                                item["dataTransforms"] = obj.get("dataTransforms")

                            if item:
                                visuals.append(item)

                        for value in obj.values():
                            walk(value)

                    elif isinstance(obj, list):
                        for item in obj:
                            walk(item)

                walk(payload)

    except Exception as exc:
        warnings.append(str(exc))
        return {"visuals": [], "warning": "Could not inspect PBIX visuals", "error": str(exc)}

    return {
        "visuals": visuals,
        "warning": None if visuals else "No visuals found in the PBIX archive",
    }

# Add a table to the Word document for the given records and columns
def add_records_table(doc, records, columns):
    # columns contains (display heading, dictionary key) pairs
    if not records:
        doc.add_paragraph("No records returned.")
        return

    table = doc.add_table(rows=1, cols=len(columns))
    table.style = "Table Grid"

    for cell, (heading, _) in zip(table.rows[0].cells, columns):
        cell.text = heading
        for run in cell.paragraphs[0].runs:
            run.bold = True

    for record in records:
        cells = table.add_row().cells
        for cell, (_, key) in zip(cells, columns):
            value = record.get(key)
            cell.text = "" if value is None else str(value)

    return table

# Create a Word document with the parsed data and user-input
def export_word_document(parsed, business_info, output_path):
    doc = Document()
    doc.add_heading(business_info["report_name"], level=0)
    
    # Add the Executive Summary section
    doc.add_heading("Executive Summary", level=1)
    p1 = doc.add_paragraph("Summary of the Power BI Report")
    p1.add_run("\n\n")
    p1.add_run(business_info["description"])
    p1.add_run("\n\nBusiness Purpose: ").bold = True
    p1.add_run(business_info["business_purpose"])
    p1.add_run("\n\nObjectives: ").bold = True
    p1.add_run(business_info["objectives"])
    p1.add_run("\n\nTarget Audience: ").bold = True
    p1.add_run(business_info["target_audience"])
    p1.add_run("\n\nScope: ").bold = True
    p1.add_run(business_info["scope"])
    p1.add_run("\n\nOut of Scope: ").bold = True
    p1.add_run(business_info["out_of_scope"])
    
    # Add the Business Information section
    doc.add_heading("Business Information", level=1)
    p2 = doc.add_paragraph("Business Information Details")
    p2.add_run("\n\n")
    p2.add_run("Business Owner: ").bold = True
    p2.add_run(business_info["business_owner"])
    p2.add_run("\nTechnical Owner: ").bold = True
    p2.add_run(business_info["technical_owner"])
    p2.add_run("\nSupport Team: ").bold = True
    p2.add_run(business_info["support_team"])
    p2.add_run("\nMaintainer: ").bold = True
    p2.add_run(business_info["maintainer"])
    
    # Add the Technical Information section
    doc.add_heading("Technical Information", level=1)
    p3 = doc.add_paragraph("Technical Information Details")
    p3.add_run("\n\n")
    p3.add_run("Refresh Schedule: ").bold = True
    p3.add_run(business_info["refresh_schedule"])
    p3.add_run("\nData Classification: ").bold = True
    p3.add_run(business_info["data_classification"])
    p3.add_run("\nDependencies: ").bold = True
    p3.add_run(business_info["dependencies"])
    p3.add_run("\nKnown Limitations: ").bold = True
    p3.add_run(business_info["known_limitations"])
    p3.add_run("\nAssumptions: ").bold = True
    p3.add_run(business_info["assumptions"])
    p3.add_run("\nRisk Assessment: ").bold = True
    p3.add_run(business_info["risk_assessment"])

    # Add the Data Sources section
    doc.add_heading("Data Sources", level=1)
    p4 = doc.add_paragraph("Data Sources Details")
    p4.add_run('bold').bold = True
    p4.add_run("\n\n")
    p4.add_run("Connections: ").bold = True
    p4.add_run("\n".join(
        json.dumps(connection, ensure_ascii=False, default=str)
        for connection in parsed["connections"]
    ))
    p4.add_run("\n\nTables: ").bold = True
    p4.add_run(", ".join(parsed["tables"]))

    # Add the Power Query section
    doc.add_heading("Power Query", level=1)
    p9 = doc.add_paragraph("Power Query Details")
    p9.add_run("\n\n")
    # Add the Power Query Records table
    p9.add_run("Power Query Records: ").bold = True
    add_records_table(doc, parsed["power_query"], [
        ("Table Name", 'TableName'), 
        ("Expression", 'Expression')
    ])
    # Add the M Parameters Records table
    p9.add_run("\n\nM Parameters: ").bold = True
    add_records_table(doc, parsed["m_parameters"], [
        ("Parameter Name", "ParameterName"),
        ("Description", "Description"),
        ("Expression", "Expression"),
        ("Modified Time", "ModifiedTime"),
    ])

    # Add the Semantic Model section
    doc.add_heading("Semantic Model", level=1)
    p5 = doc.add_paragraph("Semantic Model Details")
    p5.add_run('bold').bold = True
    p5.add_run("\n\n")
    # Add the Schema Records table
    p5.add_run("Schema: ").bold = True
    add_records_table(doc, parsed["schema"], [
        ("Table Name", "TableName"),
        ("Column Name", "ColumnName"),
        ("Pandas Data Type", "PandasDataType"),
    ])
    # Add the Measures Tables Records table
    p5.add_run("\n\nMeasures: ").bold = True
    add_records_table(doc, parsed["measures"], [
        ("Table Name", "TableName"),
        ("Measure", "Name"),
        ("Expression", "Expression"),
    ])
    # Add the DAX Columns Records table
    p5.add_run("\n\nDAX Columns: ").bold = True
    add_records_table(doc, parsed["dax_columns"], [
        ("Table Name", "TableName"),
        ("Column Name", "ColumnName"),
        ("Expression", "Expression"),
    ])

    # Add the Relationships section
    doc.add_heading("Relationships", level=1)
    p6 = doc.add_paragraph("Relationships Details")
    p6.add_run('bold').bold = True
    p6.add_run("\n\n")
    # Add the Relationships Records table
    p6.add_run("Relationships: ").bold = True
    add_records_table(doc, parsed["relationships"], [
        ("From Table Name", "FromTableName"),
        ("From Column Name", "FromColumnName"),
        ("To Table Name", "ToTableName"),
        ("To Column Name", "ToColumnName"),
        ("Is Active", "IsActive"),
        ("Cardinality", "Cardinality"),
        ("Cross Filter Behavior", "CrossFilteringBehavior"),
        ("From Key Count", "FromKeyCount"),
        ("To Key Count", "ToKeyCount"),
        ("Rely On Referential Integrity", "RelyOnReferentialIntegrity")
    ])

    # Add the Aggregations section
    doc.add_heading("Aggregations", level=1)
    p10 = doc.add_paragraph("Aggregations Details")
    p10.add_run('bold').bold = True
    p10.add_run("\n\n")
    # Add the Aggregations Records table
    p10.add_run("Aggregations: ").bold = True
    add_records_table(doc, parsed["aggregations"], [
        ("Aggregation Table", "AggregationTable"),
        ("Aggregation Column", "AggregationColumn"),
        ("Summarization", "Summarization"),
        ("Detail Table", "DetailTable"),
        ("Detail Column", "DetailColumn")
    ])

    # Add the Row Level Security (RLS) section
    doc.add_heading("Row Level Security (RLS)", level=1)
    p8 = doc.add_paragraph("Row Level Security (RLS) Details")
    p8.add_run('bold').bold = True
    p8.add_run("\n\n")
    # Add the Row Level Security (RLS) Records table
    p8.add_run("Row Level Security (RLS): ").bold = True
    add_records_table(doc, parsed["rls"], [
        ("Table Name", "TableName"),
        ("Role Name", "RoleName"),
        ("Role Description", "RoleDescription"),
        ("Filter Expression", "FilterExpression"),
        ("State", "State"),
        ("Metadata Permissions", "MetadataPermission")
    ])
        
    # Add the Visualizations section
    doc.add_heading("Visualizations", level=1)
    p7 = doc.add_paragraph("Visualizations Details")
    p7.add_run('bold').bold = True
    p7.add_run("\n\n")
    for visual in parsed.get("visuals", []):
        p7.add_run(
            f"Name: {visual.get('name')} | "
            f"Type: {visual.get('visualType')} | "
            f"Layout: {visual.get('layout')}\n"
        )
        for f in visual.get("filters", []) or []:
            p7.add_run(f"Filter: {f}\n")
        for c in visual.get("columns", []) or []:
            p7.add_run(f"Column: {c}\n")

    doc.save(output_path)

# Download button
def download_word_document(output_path):
    with open(output_path, "rb") as f:
        st.download_button(
            label="Download Word Document",
            data=f.read(),
            file_name="powerbi_documentation.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

# Create the main title for the app
st.title("PowerBI Documentation Generator")
st.subheader("Generate comprehensive documentation for your PowerBI reports with ease.")

# Form Container
with st.form(key="user_info_form"):
    st.subheader("Please fill out your details")

    # Current Power BI Document Upload Section (to autofill fields if possible)
    Current_document = st.file_uploader(
        label="Upload Current Power BI Document (Optional)",
        type=["Word", "docx"],
        help="Upload a Word document to autofill fields if available."
    )

    # Add the file uploader constrained to .zip files
    uploaded_zip = st.file_uploader(
        label="Select .pbix file to upload",
        type=["pbix"],
        help="Only .pbix files are supported."
    )

    #===============================================
    # Section 1
    #===============================================
    st.markdown("### 1. Ownership Information")
    st.divider() # Optional visual line divider

    col1, col2 = st.columns(2)
    # inputs inside the form
    with col1:
        business_owner = st.text_input("Business Owner Name")
        technical_owner = st.text_input("Technical Owner Name")
    with col2:
        support_team = st.text_input("Support team")
        maintainer = st.text_input("Maintainer")
    
    # Add a blank line for spacing
    st.write("")

    #===============================================
    # Section 2
    #===============================================
    st.markdown("### 2. Report Information")
    st.divider() # Optional visual line divider

    report_name = st.text_input("Report Name")

    col1, col2 = st.columns(2)
    with col1:
        description = st.text_area("Report Description", height=150)
        business_purpose = st.text_area("Business Purpose", height=150)
        objectives = st.text_area("Objectives", height=150)
    with col2:
        target_audience = st.text_area("Target Audience", height=150)
        scope = st.text_area("Scope", height=150)
        out_of_scope = st.text_area("Out of Scope", height=150)

    # Add a blank line for spacing
    st.write("")

    #===============================================
    # Section 3
    #===============================================
    st.markdown("### 3. Operational Information")
    st.divider() # Optional visual line divider

    col1, col2 = st.columns(2)
    with col1:
        refresh_schedule = st.text_input("Refresh Schedule")
        data_classification = st.text_input("Data Classification")
        dependencies = st.text_area("Dependencies", height=150)
    with col2:
        known_limitations = st.text_area("Known Limitations", height=150)
        assumptions = st.text_area("Assumptions", height=150)
        risk_assessment = st.text_area("Risk Assessment", height=150)
    
    # Add submit button
    submit_button = st.form_submit_button(label="Submit Form")

# Submit button logic
if submit_button:
    required_fields = {
        "Business Owner": business_owner,
        "Technical Owner": technical_owner,
        "Report Name": report_name,
        "Description": description,
        "Business Purpose": business_purpose,
        "Objectives": objectives,
        "Target Audience": target_audience,
        "Scope": scope,
        "Out of Scope": out_of_scope,
    }

    missing = [name for name, value in required_fields.items() if value.strip() == ""]

    if missing:
        st.error("Please fill in all required fields: " + ", ".join(missing) + ".")
    elif uploaded_zip is None:
        st.error("Please select a Power BI .pbix file before clicking submit.")
    elif not uploaded_zip.name.lower().endswith(".pbix"):
        st.error("Only .pbix files are supported.")
    else:
        parsed = parse_pbixray(uploaded_zip)
        visuals_info = extract_visuals_from_pbix(uploaded_zip)
        parsed["visuals"] = visuals_info["visuals"]
        if visuals_info.get("warning"):
            st.warning(visuals_info["warning"])

        business_info = {
            "business_owner": business_owner,
            "technical_owner": technical_owner,
            "support_team": support_team,
            "maintainer": maintainer,
            "report_name": report_name,
            "description": description,
            "business_purpose": business_purpose,
            "objectives": objectives,
            "target_audience": target_audience,
            "scope": scope,
            "out_of_scope": out_of_scope,
            "refresh_schedule": refresh_schedule,
            "data_classification": data_classification,
            "dependencies": dependencies,
            "known_limitations": known_limitations,
            "assumptions": assumptions,
            "risk_assessment": risk_assessment,
        }

        output_dir = Path("output")
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / "powerbi_documentation.docx"

        export_word_document(parsed, business_info, str(output_path))
        st.success("Documentation package generated successfully.")

        with open(output_path, "rb") as f:
            st.download_button(
                label="Download Word Document",
                data=f.read(),
                file_name="powerbi_documentation.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
