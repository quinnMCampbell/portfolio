# Best Practice Search System

This project reformed and streamlined access to the company's **Best Practice documents** — a collection of around 200 Word files containing procedures and policies for various floor operations.  

Previously, workers had to manually navigate a file server of PDFs to find the needed document. The new system provides a **searchable interface** that links directly to the relevant PDFs, minimizing the effort required to access information.

## Project Goals
- Index and reformat the collection of documents  
- Enable document search by **Title**, **Content**, or **Engineering-defined Keywords**  
- Provide a simple way for the engineering department to **update and maintain** the system  

## Key Contributions
- Built a **searchable HTML-based system** integrated into the company intranet  
- Designed a **user-friendly UI** with a familiar search-engine style experience  
- Organized Word documents by ID and created an **Excel index** for maintainability, including:
  - Metadata and keywords for each document  
  - Hyperlinks to original Word files for engineering access  
- Developed **PowerShell scripts** to:  
  - Update individual documents by ID  
  - Perform a mass keyword update across all documents  
- Authored a **Best Practice** documenting how to use and maintain the system  
- Conducted **early floor testing** with select departments to validate functionality, gather feedback, and ensure a smooth transition between systems  

## Technical Details
- The search query is parsed to isolate relevant terms and find matches  
- Results are ranked by relevance using an **AVL Binary Search Tree**  
- Document metadata is stored within the HTML file, due to server limitations and the low volume of documents  