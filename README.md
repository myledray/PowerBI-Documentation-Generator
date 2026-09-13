# PowerBI-Documentation-Generator
A lightweight, scalable application for generating documentation for Power BI projects (.pbix files only).

**Project Status:** Prototype / In development

## Overview 

Power BI Documentation Generator aims to reduce the manual effor involed in documenting Power BI projects.

### Supported Capabilities 

| Area | Current Support |
| --- | --- |
| Input formats | [Supported formats] |
| Documented metadata | [Supported model or report elements] |
| Output formats | [Supported formats] |
| Execution | [CLI / Local application / Web application] |

### Scope and limitations

[Describe the boundaries users should understand before getting started.]

- [Current limitation]
- [Unsupported input or feature]
- [Any prerequisites imposed by the supported workflow]

## Example output

[Include a short excerpt or screenshot from actual generated documentation.
Identify the input used and omit sensitive information.]

## Getting started

### Prerequisites

- [Required runtime and supported version]
- [Required dependencies or tools]
- [Power BI project preparation, if needed]
- [Credentials or permissions, if needed]

### Installation

[Insert the exact commands needed to install and configure the application
from a clean checkout.]

### Generate your first document

1. [Prepare a supported input.]
2. [Start the application or run the generation command.]
3. [Select or specify the input and output location.]
4. [Generate and open the documentation.]

**Expected result:** [Describe the generated files and where to find them.]

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

## What the documentation Includes 
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
