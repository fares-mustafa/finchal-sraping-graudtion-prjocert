import argparse
import sys
from logic import setup_driver, get_links, get_financial_data, download_pdfs, convert_pdf_to_excel

def main():
    parser = argparse.ArgumentParser(description='Financial Data Scraper CLI')
    parser.add_argument('-S', '--stocks', type=str, required=True, help='Comma separated list of stock names')
    parser.add_argument('-O', '--operation', type=str, choices=['F.s', 'pdf', 'convert'], required=True, help='Operation to perform: F.s for financial data, pdf for PDF download, convert for PDF to Excel conversion')
    parser.add_argument('-o', '--output', type=str, required=False, help='Output filename for CSV data or output directory for conversion')
    args = parser.parse_args()

    driver = setup_driver()
    names_links_dict = get_links(driver)
    selected_companies = args.stocks.split(',')

    if args.operation == 'F.s':
        if not args.output:
            print("Output filename is required for financial data operation")
            sys.exit(1)
        get_financial_data(driver, names_links_dict, selected_companies, args.output)

    elif args.operation == 'pdf':
        download_pdfs(driver, names_links_dict, selected_companies)

    elif args.operation == 'convert':
        if not args.output:
            print("Output directory is required for conversion operation")
            sys.exit(1)
        pdf_files = [(f"{company}_pdf_{i+1}.pdf", company, 'unknown') for i, company in enumerate(selected_companies)]
        convert_pdf_to_excel(pdf_files, args.output)

if __name__ == "__main__":
    main()

