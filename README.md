# Healthcare Document Automation & Data Extraction

## Overview

This project automated a healthcare document workflow by extracting structured data from assessment PDFs, organizing the information in Excel, and generating completed service planning documents through an automated Mail Merge process.

The solution was built to reduce manual data entry, improve document accuracy, preserve formatting, and streamline high-volume administrative workflows.

## Business Problem

Staff were required to manually review assessment PDFs, transfer data and generate individualized documents, format each document, and prepare files for distribution.

This process was time-consuming, repetitive, and prone to formatting or data-entry errors.

## Solution

I developed an automation workflow using Python, VBA, Excel, and Mail Merge to:

- Extract structured data from assessment PDFs
- Organize extracted fields into Excel
- Generate individualized service planning documents
- Preserve document formatting
- Keep required text sections editable
- Reduce manual document preparation time

## Tools Used

- Python
- VBA
- Microsoft Excel
- Microsoft Word Mail Merge
- PDF data extraction
- RingCentral workflow support

## Business Impact

- Automated 100+ patient-specific documents per day
- Saved approximately 40–50 labor hours per week
- Reduced manual data entry and document formatting work
- Improved document consistency and accuracy
- Created documentation shared across the office for training and scalability

## Workflow

Assessment PDF  
↓  
Python Data Extraction  
↓  
Structured Excel Dataset  
↓  
VBA / Mail Merge Automation  
↓  
Formatted Service Planning Documents  
↓  
Review and Distribution

## Repository Structure

```text
healthcare-document-automation/

├── docs/
│   └── data_privacy_note.md
│   └── workflow_overview.md
│
├── samples/
│   └── sample_assesment_data.xlsx
│   └── sample_output_description.md
│
├── src/
│   └── document_automation.py
│   └── locked_merged_docs_macro.bas
│   └── medication_table_macro.bas
│
├── README.md
├── LICENSE
├── .gitignore

    

