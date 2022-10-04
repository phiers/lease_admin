#### Overview of Lease Admin Billing Process

### Setting Up
1. Required Folders
2. Copy files emailed from Lx into "1_Lx_files" folder (will replace with DB)
3. Setup inputs
    a. Last month's intial invoice analysis - as cleaned (cleaned means it matches exactly the final output) - becomes this month's lm_invoice_analysis, which must be added to the "4_input_files" folder. 
    b. Any price changes made on 'type_desc_price_matrix.csv' file
    c. New customer? Add to customer_names.csv
    d. Fill out additional_invoice_items.csv for any items (abstracts, implementation fees, etc.)
    e. The update of teh type_desc_price_matrix.csv file may need to happen again for any new lxcodes (e.g., a client has never had a lease in Future Possession status - the code and price for that will  need to be added)
4. Run process __ to read Lx files and create lease_files
5. Run process __ to create intial_invoice_analysis
    a. This is a working document to create the final analysis.
    b. Clean up and make sure it is 100% correct, including recalculating rows for totals (price x qnty). The current month matters as it will be used next month as last month; the last month doesn't get used again
6. Once the initial analysis is complete and correct, run final analysis
