#!/usr/bin/env python3
"""
Email Processing Script
Filters input.txt based on send_log.txt status and creates unsent.txt
"""

def parse_send_log(log_file):
    """
    Parse send_log.txt and return a dict with email status.
    Returns: dict {email: status} where status is 'SENT' or 'FAILED'
    """
    email_status = {}
    
    try:
        with open(log_file, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                
                # Extract the "To:" email address
                if '| To: ' in line:
                    to_part = line.split('| To: ')[1].split(' |')[0].strip()
                    
                    # Extract the status
                    if '| Status: ' in line:
                        status_part = line.split('| Status: ')[1].split(' |')[0].strip()
                        email_status[to_part] = status_part
                
    except FileNotFoundError:
        print(f"Warning: {log_file} not found. Treating all emails as unsent.")
    except Exception as e:
        print(f"Error parsing send_log.txt: {e}")
    
    return email_status


def process_input_file(input_file, send_log_file, output_file):
    """
    Process input.txt and create unsent.txt based on send_log status.
    - Remove lines where the first email has status SENT
    - Keep lines where the first email has status FAILED or is not in send_log
    """
    email_status = parse_send_log(send_log_file)
    unsent_lines = []
    
    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                
                # Extract the first email (before the first |)
                first_email = line.split('|')[0].strip()
                
                # Check status
                status = email_status.get(first_email, 'NOT_SENT')
                
                if status == 'SENT':
                    # Skip this line - it was successfully sent
                    print(f"Skipping (SENT): {first_email}")
                else:
                    # Keep this line - it failed or wasn't sent
                    unsent_lines.append(line)
                    if status == 'FAILED':
                        print(f"Adding (FAILED): {first_email}")
                    else:
                        print(f"Adding (NOT SENT): {first_email}")
        
        # Write unsent emails to output file
        with open(output_file, 'w', encoding='utf-8') as f:
            for line in unsent_lines:
                f.write(line + '\n')
        
        print(f"\nProcessing complete!")
        print(f"Total unsent emails: {len(unsent_lines)}")
        print(f"Output written to: {output_file}")
        
    except FileNotFoundError:
        print(f"Error: {input_file} not found.")
    except Exception as e:
        print(f"Error processing files: {e}")


if __name__ == "__main__":
    # File paths
    INPUT_FILE = "input.txt"
    SEND_LOG_FILE = "send_log.txt"
    OUTPUT_FILE = "unsent.txt"
    
    print("Starting email processing...")
    print(f"Input file: {INPUT_FILE}")
    print(f"Send log file: {SEND_LOG_FILE}")
    print(f"Output file: {OUTPUT_FILE}\n")
    
    process_input_file(INPUT_FILE, SEND_LOG_FILE, OUTPUT_FILE)
