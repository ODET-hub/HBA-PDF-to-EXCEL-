#!/usr/bin/env python3
"""
Script to test the improved bank statement processing.
"""

import requests
import os

def test_bank_statement():
    """Test uploading a bank statement PDF to the converter."""
    
    # Check if bank statement PDF exists
    pdf_path = "uploads/Bank_Statement_Domiciliary_Account_June_1-Sep_29_2024.pdf"
    if not os.path.exists(pdf_path):
        print(f"Error: {pdf_path} not found.")
        return
    
    # URL of the Flask app
    url = "http://127.0.0.1:5000/"
    
    # Prepare the file upload
    with open(pdf_path, 'rb') as f:
        files = {'file': (os.path.basename(pdf_path), f, 'application/pdf')}
        
        print(f"Uploading bank statement to {url}...")
        
        try:
            # Upload the file
            response = requests.post(url, files=files, allow_redirects=False)
            
            if response.status_code == 302:  # Redirect expected
                # Get the redirect location
                redirect_url = response.headers.get('Location')
                if redirect_url:
                    print(f"Upload successful! Redirecting to: {redirect_url}")
                    
                    # Follow the redirect to download the file
                    download_response = requests.get(f"http://127.0.0.1:5000{redirect_url}")
                    
                    if download_response.status_code == 200:
                        # Save the Excel file
                        excel_filename = redirect_url.split('/')[-1]
                        with open(excel_filename, 'wb') as excel_file:
                            excel_file.write(download_response.content)
                        print(f"Excel file saved as: {excel_filename}")
                        print("✅ Bank statement processed successfully!")
                    else:
                        print(f"Download failed with status: {download_response.status_code}")
                else:
                    print("Upload successful but no redirect URL found")
            else:
                print(f"Upload failed with status: {response.status_code}")
                print(f"Response: {response.text}")
                
        except requests.exceptions.ConnectionError:
            print("Error: Could not connect to the Flask app. Make sure it's running on http://127.0.0.1:5000")
        except Exception as e:
            print(f"Error: {e}")

if __name__ == "__main__":
    test_bank_statement()
