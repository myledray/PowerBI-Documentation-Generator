# PowerBI-Documentation-Generator
An app built to generate extensive documentation by uploading .pbix files

## Architecture 
#### StreamLit UI
- Business Metadata Form
- Business Metadata Form
- Submit Button

#### Python Processing Layer 
- parse_pbixray 
- extract_visuals_from_pbix
- export_word_document
- download_word_document
 
## Requirements 
#### Project Vision 
1. Collect Business documentation from users 
2. Extracts technical documentation from a Power BI file 
3. Combines both into a professional documentation package 
4. Exports the result as a word document 

## Docuementation Generated Sections 
1. Heading
    - Report Name
2. Executive Summary
    - Description
    - Business Purpose
    - Objectives
    - Target Audience
    - Scope
    - Out of Scope
3. Business Information
    - Business Owner
    - Technical Owner
    - Maintainer
4. Technical Information 
    - Refresh Schedule
    - Data Classification
    - Dependencies
    - Known Limitations 
    - Assumptions
    - Risk Assessment
5. Data Sources 
    - Connections
    - Tables
6. Power Query
    - Power Query
        - Table Name 
        - Expressions 
    - M Parameters
        - Parameter Name 
        - Description
        - Expression
        - Modified Time
7. Semantic Model 
    - Schema
        - Table Name
        - Column Name
        - Pandas Data type
    - Measures
        - Table Name
        - Meassure
        - Expression
    - DAX Columns
        - Table Name
        - Column Name
        - Expression
8. Relationships
    - Relationships
        - From Table Name
        - Frome Column Name
        - To Table Name
        - To Column Name
        - Is Active
        - Cardinality
        - Cross Filter Behavior
        - From Key Count 
        - To Key Count
        - Rely On Referencial Integrity
9. Aggregations
    - Aggregations 
        - Aggregation Table 
        - Aggregation Column
        - Summarization
        - Detail Table
        - Detail Column
10. Row Level Security (RLS)
    - rls
        - Table Name
        - Role Name 
        - Role Description
        - Filter Expression
        - State
        - Metadata Permissions
11. Visualizations
    - Visuals 
        - Name 
        - Type 
        - Layout
