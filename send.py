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
STAGGER_DELAY = 30  # seconds between instance starts

# === ORIGINAL CONFIG ===
LOGINS_FILE = "logins.txt"           
RECIPIENTS_FILE = "input.txt"        
STATIC_ATTACHMENT = "W9.pdf"
BASE_SUBJECT = "Re: {company_name} - Overdue Strategic Planning Services Invoice INV172573"
LOG_FILE = "send_log.txt"
EMAILS_PER_ACCOUNT = 4
HEADLESS = False  # Set to True for production
REPLY_TO = ""
CC_EMAIL = "Vincent Corwin <vincent@quantproviders.com>"
BCC_EMAIL = ""

# PDF Generation Config
WKHTMLTOPDF_PATH = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
ENABLE_PDF_GENERATION = True
PDF_CLEANUP = True

# Template Config
MESSAGES_FOLDER = "messages"
HTML_TEMPLATE_FILE = "template.html"
BOX_DOMAIN = "quantproviders.com"

# Attachment Config
ENABLE_ATTACHMENTS = True

# Delete functionality config
MAX_DELETE_RETRIES = 3  # Increased from 2 for critical delete operations
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

# === MULTI-INSTANCE SETUP ===
def split_data_in_memory():
    """Split data in memory without creating files"""
    print("📂 Loading and splitting data in memory...")
    
    # Load original files
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
    
    # Calculate splits
    logins_per_instance = max(1, len(all_logins) // NUM_INSTANCES)
    recipients_per_instance = max(1, len(all_recipients) // NUM_INSTANCES)
    
    print(f"📊 Splitting {len(all_logins)} accounts and {len(all_recipients)} recipients")
    print(f"📊 Each instance: ~{logins_per_instance} accounts, ~{recipients_per_instance} recipients")
    
    # Create instance data
    instance_data = []
    for i in range(NUM_INSTANCES):
        instance_num = i + 1
        
        # Split logins
        start_login = i * logins_per_instance
        end_login = len(all_logins) if i == NUM_INSTANCES - 1 else start_login + logins_per_instance
        instance_logins = all_logins[start_login:end_login]
        
        # Split recipients
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
    """Clean up temporary instance files"""
    print("🧹 Cleaning up temporary files...")
    temp_dir = Path("tdata")
    
    if temp_dir.exists():
        try:
            # Remove all files in tdata directory
            for file_path in temp_dir.iterdir():
                try:
                    if file_path.is_file():
                        file_path.unlink()
                        print(f"🗑️ Removed {file_path}")
                except Exception as e:
                    print(f"[!] Could not remove {file_path}: {e}")
            
            # Remove the directory itself
            temp_dir.rmdir()
            print(f"🗑️ Removed tdata directory")
        except Exception as e:
            print(f"[!] Could not remove tdata directory: {e}")
    
    print("✅ Cleanup completed!")

def monitor_instances():
    """Monitor all running instances and display live stats"""
    print("📊 Starting live monitoring...")
    
    while True:
        # Clear screen
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
        
        # Check shared log file
        if Path(LOG_FILE).exists():
            try:
                with open(LOG_FILE, 'r', encoding='utf-8') as f:
                    content = f.read()
                    
                # Count total emails
                total_sent = content.count("Status: SENT")
                total_failed = content.count("Status: FAILED") 
                total_attempted = total_sent + total_failed
                
                # Count by instance (look for Instance markers)
                for i in range(1, NUM_INSTANCES + 1):
                    instance_sent = content.count(f"[Instance {i}]") and content.count("Status: SENT")
                    instance_failed = content.count(f"[Instance {i}]") and content.count("Status: FAILED")
                    
                    # Simple count for each instance (approximate)
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
        
        # Display instance stats
        for i in range(1, NUM_INSTANCES + 1):
            stats = instance_stats.get(i, {'sent': 0, 'failed': 0, 'attempted': 0, 'status': 'NOT STARTED'})
            print(f"📧 Instance {i}: {stats['status']:<12} | Sent: {stats['sent']:<3} | Failed: {stats['failed']:<3} | Total: {stats['attempted']:<3}")
        
        print("-" * 80)
        success_rate = (total_sent / total_attempted * 100) if total_attempted > 0 else 0
        print(f"📊 OVERALL: Sent: {total_sent:<3} | Failed: {total_failed:<3} | Total: {total_attempted:<3} | Rate: {success_rate:.1f}%")
        print(f"🔄 Instances Running: {instances_running}/{NUM_INSTANCES}")
        print("="*80)
        
        # Check if all completed (when no more activity for a while)
        if instances_running == 0 and total_attempted > 0:
            print("\n🎉 ALL INSTANCES COMPLETED!")
            
            # Generate final report
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
        
        # Wait before next update
        time.sleep(5)

# === ALL YOUR ORIGINAL FUNCTIONS (unchanged) ===

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
            
            # Verify it's actually a PDF by checking file header
            with open(pdf_path, 'rb') as f:
                header = f.read(4)
                if header == b'%PDF':
                    print(f"[✓] PDF file validated successfully")
                    return str(pdf_path.resolve())  # Return absolute path
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
        # Method 1: Direct file input (faster and more reliable)
        print("[>] Trying direct file input method...")
        file_inputs = page.locator("input[type='file']")
        input_count = file_inputs.count()
        
        if input_count > 0:
            print(f"[>] Found {input_count} file inputs, trying each...")
            
            for i in range(input_count):
                try:
                    file_inputs.nth(i).set_input_files(abs_path)
                    
                    # Reduced wait time for faster operation
                    wait_time = 2000 if FAST_MODE else 3000
                    page.wait_for_timeout(wait_time)
                    
                    # Quick verification
                    if page.locator(f"text='{file_name}'").is_visible():
                        print(f"[✓] {description} attached via file input #{i+1}")
                        return True
                        
                except Exception as e:
                    print(f"[!] File input #{i+1} failed: {e}")
                    continue
        
        # Method 2: Use expect_file_chooser (fallback)
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
                    
                    # Use proper file chooser handling
                    with page.expect_file_chooser() as fc_info:
                        attach_btn.click()
                    
                    file_chooser = fc_info.value
                    file_chooser.set_files(abs_path)
                    
                    # Reduced wait and verify attachment
                    wait_time = 3000 if FAST_MODE else 5000
                    page.wait_for_timeout(wait_time)
                    
                    # Quick verification
                    if page.locator(f"text='{file_name}'").is_visible():
                        print(f"[✓] {description} verified as attached in UI")
                        return True
                    
                    # Check for attachment area/container
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
        # Wait for any popups to appear
        page.wait_for_timeout(1500 if FAST_MODE else 2500)
        
        # Check for file type error first
        if page.locator("text=\"The following files weren't inserted because they aren't supported image file types\"").is_visible():
            print("[!] CRITICAL: Outlook rejected file as unsupported image type")
            try:
                page.click("button[aria-label='Close'], button:has-text('×'), button:has-text('OK')")
                print("[>] Closed file type error popup")
            except:
                page.keyboard.press("Escape")
            return "FILE_TYPE_ERROR"
        
        # Check for attachment reminder
        if page.locator("text='Attachment reminder'").is_visible():
            print("[>] Attachment reminder popup appeared")
            
            # Quick check if we actually have attachments
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
        
        # Check for other common popups/overlays
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
                    # Try to close it
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
                        # If no close button found, try Escape
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

def load_logins():
    """Load login credentials from file"""
    logins = []
    with open(LOGINS_FILE, 'r') as f:
        for line in f:
            line = line.strip()
            if line.count(":") >= 1:
                parts = line.split(":", 2)
                if len(parts) == 2:
                    logins.append([parts[0], parts[1], ""])
                else:
                    logins.append(parts)
    return logins

def load_recipients():
    """Load recipient data from file"""
    recipients = []
    
    try:
        with open(RECIPIENTS_FILE, 'r', encoding='utf-8') as f:
            for line_num, line in enumerate(f, 1):
                line = line.strip()
                if not line or line.startswith('#'):
                    continue
                
                if ' | ' in line:
                    parts = [part.strip() for part in line.split(' | ')]
                    if len(parts) >= 6:
                        email, name, company_name, domain, username, ceo = parts[:6]
                        address_raw = parts[6] if len(parts) >= 7 else "United States"
                        
                        recipients.append({
                            'email': email,
                            'name': name,
                            'company_name': company_name,
                            'domain': domain,
                            'username': username,
                            'ceo': ceo,
                            'address_raw': address_raw
                        })
                    else:
                        print(f"[!] Line {line_num}: Insufficient data")
                        continue
                        
                elif "@" in line:
                    email = line
                    username = email.split('@')[0]
                    domain = email.split('@')[1]
                    company_name = domain.split('.')[0].title()
                    
                    recipients.append({
                        'email': email,
                        'name': username.title(),
                        'company_name': company_name,
                        'domain': domain,
                        'username': username,
                        'ceo': 'CEO',
                        'address_raw': 'United States'
                    })
                    
    except FileNotFoundError:
        print(f"[!] Recipients file {RECIPIENTS_FILE} not found")
        return []
    except Exception as e:
        print(f"[!] Error loading recipients: {e}")
        return []
    
    print(f"[+] Loaded {len(recipients)} recipients")
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
    """Write detailed sending results to shared log file"""
    instance_num = os.environ.get('INSTANCE_NUM', '1')
    
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        timestamp = datetime.now().isoformat()
        company = recipient_data.get('company_name', 'Unknown')
        f.write(f"{timestamp} | [Instance {instance_num}] From: {sender} | To: {recipient_data['email']} | Company: {company} | Status: {status} | {attachment_info}\n")

def detect_sender_name(page):
    """Optimized sender name detection with faster timeouts"""
    try:
        print("[>] Detecting sender name...")
        # Reduced wait time
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
                # Reduced wait time
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
        
        # Reduced wait time
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
        
        # Faster timeout for password field
        page.wait_for_selector("input[type='password']", timeout=TIMEOUT_FAST)
        page.fill("input[type='password']", password)
        page.click("input[type='submit']")
        
        # Reduced wait time
        page.wait_for_timeout(3000 if FAST_MODE else 5000)

        try:
            if page.locator("input[id='idBtn_Back']").is_visible():
                page.click("input[id='idBtn_Back']")
        except:
            pass

        # Reduced wait time
        page.wait_for_timeout(5000 if FAST_MODE else 8000)
        
        # Faster timeout for reaching interface
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

def delete_last_sent_email_robust(page):
    """Enhanced delete function to delete ONLY the most recent email (first in list)"""
    for attempt in range(MAX_DELETE_RETRIES):
        try:
            print(f"[>] Delete attempt {attempt + 1}/{MAX_DELETE_RETRIES}")
            
            if attempt > 0:
                page.wait_for_timeout(DELETE_RETRY_DELAY * 1000)
            
            # Navigate to sent items with more specific URL
            print("[>] Navigating to Sent Items...")
            page.goto("https://outlook.office.com/mail/sentitems", wait_until="networkidle", timeout=15000)
            
            # Wait for sent items to fully load
            page.wait_for_timeout(3000)
            
            # Count emails before deletion for verification
            try:
                email_count_before = 0
                email_list_selectors = [
                    "[role='listbox'] [role='option']",
                    "[data-testid='mail-list'] div[role='listitem']",
                    ".ms-List .ms-List-cell",
                    "[role='grid'] [role='row']"
                ]
                
                for selector in email_list_selectors:
                    try:
                        emails = page.locator(selector)
                        if emails.count() > 0:
                            email_count_before = emails.count()
                            print(f"[>] Found {email_count_before} emails in sent items")
                            break
                    except:
                        continue
                
                if email_count_before == 0:
                    print("[!] No emails found in sent items - nothing to delete")
                    return True  # Nothing to delete is success
                    
            except Exception as e:
                print(f"[!] Could not count emails before deletion: {e}")
                email_count_before = 0
            
            # METHOD 1: Click on first email then delete
            try:
                print("[>] Method 1: Clicking first email to select it...")
                
                # Find and click the first email in the list
                first_email_selected = False
                for selector in email_list_selectors:
                    try:
                        first_email = page.locator(selector).first
                        if first_email.is_visible():
                            print(f"[>] Found first email with selector: {selector}")
                            first_email.click()
                            page.wait_for_timeout(500)  # Wait for selection
                            first_email_selected = True
                            print("[✓] First email selected")
                            break
                    except Exception as e:
                        print(f"[!] Selector {selector} failed: {e}")
                        continue
                
                if not first_email_selected:
                    print("[!] Could not select first email, trying keyboard method...")
                    # Fallback to keyboard selection
                    page.keyboard.press("Home")  # Go to first item
                    page.wait_for_timeout(500)
                    print("[>] Used Home key to select first email")
                
                # Now delete the selected email
                print("[>] Deleting selected email...")
                page.keyboard.press("Delete")
                page.wait_for_timeout(1500)  # Wait for delete action
                
                # Handle any confirmation dialogs - but be specific about NOT deleting all
                try:
                    # Check for the dangerous "Delete all" dialog first
                    if page.locator("text*=move all the conversations").is_visible() or page.locator("button:has-text('Delete all')").is_visible():
                        print("[!] DANGER: Outlook wants to delete ALL emails - clicking Cancel")
                        cancel_selectors = [
                            "button:has-text('Cancel')",
                            "button[aria-label='Cancel']"
                        ]
                        
                        cancelled = False
                        for cancel_selector in cancel_selectors:
                            if page.locator(cancel_selector).is_visible():
                                page.click(cancel_selector)
                                print("[✓] Clicked Cancel to avoid deleting all emails")
                                cancelled = True
                                break
                        
                        if not cancelled:
                            page.keyboard.press("Escape")
                            print("[✓] Used Escape to cancel delete all")
                        
                        # This method failed, continue to next method
                        continue
                    
                    # Check for single email delete confirmation
                    single_delete_confirmations = [
                        "button:has-text('Delete')",
                        "button:has-text('Yes')", 
                        "button:has-text('Move to Deleted Items')"
                    ]
                    
                    confirmation_found = False
                    for conf_selector in single_delete_confirmations:
                        if page.locator(conf_selector).is_visible():
                            # Double check this is NOT a "delete all" dialog
                            dialog_text = page.locator("[role='dialog']").text_content() if page.locator("[role='dialog']").is_visible() else ""
                            if "all" not in dialog_text.lower() and "conversations" not in dialog_text.lower():
                                print(f"[>] Found single delete confirmation, clicking: {conf_selector}")
                                page.click(conf_selector)
                                page.wait_for_timeout(1000)
                                confirmation_found = True
                                break
                            else:
                                print("[!] This appears to be a delete all dialog, cancelling...")
                                page.keyboard.press("Escape")
                                continue
                    
                    if not confirmation_found:
                        # Try Enter for simple confirmations
                        page.keyboard.press("Enter")
                        page.wait_for_timeout(1000)
                        print("[>] Used Enter key for confirmation")
                        
                except Exception as conf_error:
                    print(f"[!] Confirmation handling error: {conf_error}")
                
                # VERIFY METHOD 1 SUCCESS
                print("[>] Verifying Method 1 deletion...")
                page.wait_for_timeout(2000)  # Give time for UI to update
                
                verification_success = False
                
                # Count emails after deletion
                if email_count_before > 0:
                    try:
                        for selector in email_list_selectors:
                            try:
                                emails_after = page.locator(selector)
                                email_count_after = emails_after.count()
                                print(f"[>] Emails after deletion: {email_count_after} (was {email_count_before})")
                                
                                if email_count_after < email_count_before:
                                    print("[✓] Method 1 SUCCESS: Email count decreased by", email_count_before - email_count_after)
                                    verification_success = True
                                    break
                                elif email_count_after == 0 and email_count_before > 0:
                                    print("[✓] Method 1 SUCCESS: Sent folder is now empty")
                                    verification_success = True
                                    break
                            except:
                                continue
                    except Exception as count_error:
                        print(f"[!] Count verification failed: {count_error}")
                
                # Check for empty folder
                if not verification_success:
                    try:
                        empty_indicators = [
                            "text*=No items",
                            "text*=empty",
                            "text*=No messages"
                        ]
                        
                        for indicator in empty_indicators:
                            if page.locator(indicator).is_visible():
                                print("[✓] Method 1 SUCCESS: Empty folder detected")
                                verification_success = True
                                break
                    except:
                        pass
                
                if verification_success:
                    print("[✓] Method 1 DELETE VERIFIED SUCCESSFUL - Only one email deleted")
                    return True
                else:
                    print("[!] Method 1 verification failed - trying next method")
                    
            except Exception as e:
                print(f"[!] Method 1 failed: {e}")
            
            # METHOD 2: Alternative selection and delete
            print("[>] Method 2: Alternative first email selection...")
            
            try:
                # Recount emails for new baseline
                current_count = 0
                for selector in email_list_selectors:
                    try:
                        emails = page.locator(selector)
                        if emails.count() > 0:
                            current_count = emails.count()
                            print(f"[>] Current email count: {current_count}")
                            break
                    except:
                        continue
                
                if current_count == 0:
                    print("[✓] No emails left to delete")
                    return True
                
                # Try different approach - use arrow keys to ensure single selection
                print("[>] Using arrow key navigation to select first email...")
                page.keyboard.press("Control+Home")  # Ensure we're at the top
                page.wait_for_timeout(500)
                page.keyboard.press("Down")  # Select first item properly
                page.wait_for_timeout(500)
                page.keyboard.press("Up")    # Back to very first
                page.wait_for_timeout(500)
                
                # Now delete
                page.keyboard.press("Delete")
                page.wait_for_timeout(1500)
                
                # Handle confirmations (avoid delete all)
                if page.locator("text*=move all the conversations").is_visible():
                    print("[!] Delete all dialog appeared - cancelling")
                    page.keyboard.press("Escape")
                    continue
                
                # Confirm single delete
                if page.locator("button:has-text('Delete')").is_visible():
                    dialog_text = page.locator("[role='dialog']").text_content() if page.locator("[role='dialog']").is_visible() else ""
                    if "all" not in dialog_text.lower():
                        page.click("button:has-text('Delete')")
                        page.wait_for_timeout(1000)
                
                # Verify Method 2
                page.wait_for_timeout(2000)
                try:
                    for selector in email_list_selectors:
                        try:
                            emails_after_m2 = page.locator(selector)
                            email_count_after_m2 = emails_after_m2.count()
                            print(f"[>] Emails after Method 2: {email_count_after_m2} (was {current_count})")
                            
                            if email_count_after_m2 < current_count:
                                print("[✓] Method 2 DELETE VERIFIED SUCCESSFUL")
                                return True
                        except:
                            continue
                except:
                    pass
                
                print("[!] Method 2 verification failed")
                
            except Exception as e:
                print(f"[!] Method 2 failed: {e}")
            
            # METHOD 3: Direct element deletion (last resort)
            print("[>] Method 3: Direct element interaction...")
            
            try:
                # Find the first email element directly and try to delete it
                for selector in email_list_selectors:
                    try:
                        first_email = page.locator(selector).first
                        if first_email.is_visible():
                            # Right-click for context menu
                            first_email.click(button="right")
                            page.wait_for_timeout(1000)
                            
                            # Look for delete in context menu
                            delete_menu_items = [
                                "text=Delete",
                                "[aria-label*='Delete']",
                                "button:has-text('Delete')"
                            ]
                            
                            for menu_item in delete_menu_items:
                                if page.locator(menu_item).is_visible():
                                    page.click(menu_item)
                                    page.wait_for_timeout(1000)
                                    print("[✓] Used context menu to delete")
                                    return True
                            
                            # If no context menu, try Escape and use keyboard
                            page.keyboard.press("Escape")
                            first_email.click()
                            page.wait_for_timeout(500)
                            page.keyboard.press("Delete")
                            page.wait_for_timeout(1500)
                            
                            if not page.locator("text*=move all the conversations").is_visible():
                                print("[✓] Method 3 keyboard delete successful")
                                return True
                            else:
                                page.keyboard.press("Escape")
                            
                            break
                    except:
                        continue
                        
            except Exception as e:
                print(f"[!] Method 3 failed: {e}")
                
        except Exception as nav_error:
            print(f"[!] Navigation error on attempt {attempt + 1}: {nav_error}")
    
    print(f"[!] All {MAX_DELETE_RETRIES} delete attempts failed")
    return SKIP_DELETE_ON_FAILURE

def send_multiple_emails(email, recipients_batch, templates, sender_name, p):
    """Optimized email sending with performance improvements and better error handling"""
    try:
        browser = p.chromium.launch(headless=HEADLESS)
        session_file = f'session_{email.replace("@", "_at_")}.json'
        context = browser.new_context(storage_state=session_file)
        page = context.new_page()
        
        # Test if session is still valid
        try:
            page.goto("https://outlook.office.com/mail", timeout=30000)
            page.wait_for_timeout(3000 if FAST_MODE else 5000)
            
            # Check if we're actually logged in
            if page.locator("input[type='email']").is_visible() or page.locator("input[type='password']").is_visible():
                print(f"[!] Session expired for {email}, account needs re-login")
                browser.close()
                return False
                
        except Exception as e:
            print(f"[!] Failed to access Outlook for {email}: {e}")
            browser.close()
            return False

        batch_success = True
        
        for i, recipient_data in enumerate(recipients_batch):
            global attempted, success_count, fail_count
            attempted += 1
            
            print(f"\n[>] Processing {recipient_data['email']} ({i+1}/{len(recipients_batch)})")
            
            try:
                address_text = recipient_data['address_raw']
                template = random.choice(templates)
                
                # Pre-generate all content to save time later
                personalized_message = personalize_content(
                    template, 
                    recipient_data['email'],
                    recipient_data['name'],
                    recipient_data['company_name'],
                    recipient_data['domain'],
                    recipient_data['username'],
                    recipient_data['ceo'],
                    address_text,
                    email,  # sender_email
                    sender_name  # sender_name
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
                    email,  # sender_email
                    sender_name  # sender_name
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
                            email,  # sender_email
                            sender_name  # sender_name
                        )
                        pdf_path = generate_pdf(personalized_html, recipient_data)
                
                # Start composing email with faster selectors
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
                            # Faster timeout
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
                        # Wait for compose window to load and verify it opened
                        page.wait_for_timeout(COMPOSE_DELAY)
                        
                        # Check if compose window actually opened
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
                            
                            # Try to close any open compose windows first
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

                # Fill recipient with multiple selector attempts
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
                    
                    # Try to close the compose window
                    try:
                        page.keyboard.press("Escape")
                    except:
                        pass
                    continue

                # Handle CC with faster operations
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

                # Fill subject - faster timeout
                page.wait_for_selector("input[placeholder='Add a subject']", timeout=TIMEOUT_FAST)
                page.fill("input[placeholder='Add a subject']", subject)

                # Fill message body with faster method
                try:
                    editor_div = page.locator("div[role='textbox'][aria-label*='Message body'][contenteditable='true']")
                    editor_div.wait_for(state="visible", timeout=TIMEOUT_FAST)
                    editor_div.evaluate("(el, html) => el.innerHTML = html", personalized_message)
                    print("[✓] Body injected")
                except Exception as e:
                    print(f"[!] Message body failed: {e}")
                    continue

                # Handle attachments with optimized method
                attachment_info = []
                if ENABLE_ATTACHMENTS:
                    if pdf_path:
                        if simple_attach_file(page, pdf_path, "Generated PDF"):
                            attachment_info.append(f"PDF: {Path(pdf_path).name}")
                    
                    if STATIC_ATTACHMENT and Path(STATIC_ATTACHMENT).exists():
                        if simple_attach_file(page, STATIC_ATTACHMENT, "Static file"):
                            attachment_info.append(f"Static: {STATIC_ATTACHMENT}")

                # Send the email with optimized popup handling
                try:
                    print("[>] Preparing to send email...")
                    
                    # First, ensure we're ready to send
                    page.wait_for_timeout(1000)
                    
                    # Method 1: Try multiple Send button selectors
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
                                
                                # Scroll to button and ensure it's in view
                                send_button.scroll_into_view_if_needed()
                                page.wait_for_timeout(500)
                                
                                # Try multiple click methods
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
                                        
                                        # Check if click worked by seeing if compose is still open
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
                    
                    # Method 2: Keyboard shortcut as backup
                    if not send_clicked:
                        print("[>] Button clicks failed, trying keyboard shortcut...")
                        try:
                            # Focus on the compose area first
                            page.locator("div[role='textbox'][aria-label*='Message body']").click()
                            page.wait_for_timeout(500)
                            
                            # Try Ctrl+Enter
                            page.keyboard.press("Control+Enter")
                            page.wait_for_timeout(3000)
                            
                            if not page.locator("input[placeholder='Add a subject']").is_visible():
                                print("[✓] Send successful with Ctrl+Enter")
                                send_clicked = True
                        except Exception as kb_error:
                            print(f"[!] Keyboard shortcut failed: {kb_error}")
                    
                    # Method 3: Last resort - find ANY clickable element with "Send" text
                    if not send_clicked:
                        print("[>] Trying last resort send methods...")
                        try:
                            # Look for any element containing "Send"
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
                    
                    # Check final result
                    if send_clicked:
                        # Wait for any popups and handle them
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
                            # Success!
                            print(f"[✓] Email successfully sent to {recipient_data['email']}")
                            success_count += 1
                            attachment_str = "; ".join(attachment_info) if attachment_info else "No attachments"
                            write_log(email, recipient_data, "SENT", attachment_str)
                            
                            # Wait before delete
                            time.sleep(5 if FAST_MODE else 8)
                            delete_last_sent_email_robust(page)
                    else:
                        # All send methods failed
                        print(f"[X] All send methods failed for {recipient_data['email']} - email remains in drafts")
                        fail_count += 1
                        write_log(email, recipient_data, "FAILED", "Unable to click Send button - all methods failed")
                    
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

                # Reduced inter-email delay
                time.sleep(INTER_EMAIL_DELAY)

            except Exception as main_error:
                print(f"[X] Error processing {recipient_data['email']}: {main_error}")
                fail_count += 1
                write_log(email, recipient_data, "FAILED", f"Processing error: {main_error}")

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
    """Run single instance (original functionality)"""
    global success_count, fail_count
    
    print("="*60)
    print("🚀 EMAIL AUTOMATION SCRIPT")
    print("="*60)
    
    logins = load_logins()
    recipients = load_recipients()
    templates = load_templates()

    if not recipients or not logins or not templates:
        print("[!] Missing required files. Exiting.")
        return

    print(f"[+] Loaded {len(logins)} accounts and {len(recipients)} recipients")
    print(f"[+] Loaded {len(templates)} message templates")
    print(f"[+] PDF Generation: {'Enabled' if ENABLE_PDF_GENERATION else 'Disabled'}")
    print(f"[+] Attachments: {'Enabled' if ENABLE_ATTACHMENTS else 'Disabled'}")

    login_index = 0
    recipient_index = 0
    logins_updated = False
    failed_logins = set()  # Track failed logins to skip them

    with sync_playwright() as p:
        while recipient_index < len(recipients):
            # Find next working login (skip failed ones)
            attempts = 0
            current_login = None
            
            while attempts < len(logins):
                email, password, sender_name = logins[login_index % len(logins)]
                
                # Skip if this login has failed before
                if email in failed_logins:
                    print(f"[!] Skipping previously failed login: {email}")
                    login_index += 1
                    attempts += 1
                    continue
                
                current_login = (email, password, sender_name)
                break
            
            # If all logins have failed, exit
            if current_login is None:
                print(f"[!] All logins have failed! Cannot continue.")
                break
                
            email, password, sender_name = current_login

            # Check session or login
            if not session_exists(email):
                print(f"[>] Attempting login for: {email}")
                login_success, detected_name = login_and_save_session(email, password, p)
                
                if not login_success:
                    print(f"[!] Login failed for {email}, marking as failed and trying next account")
                    failed_logins.add(email)
                    login_index += 1
                    continue  # Try next login without processing recipients
                
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

            # Record start time for performance metrics
            batch_start_time = time.time()
            
            # Try to send emails with this account
            batch_success = send_multiple_emails(email, batch, templates, sender_name, p)
            
            # If the entire batch failed due to account issues, mark account as failed
            if not batch_success:
                print(f"[!] Account {email} failed during sending, marking as failed")
                failed_logins.add(email)
                login_index += 1
                continue  # Don't advance recipient_index, try with next account
            
            # Calculate and display performance metrics
            batch_time = time.time() - batch_start_time
            emails_per_minute = (len(batch) / batch_time) * 60 if batch_time > 0 else 0
            print(f"[📊] Batch completed in {batch_time:.1f}s ({emails_per_minute:.1f} emails/min)")

            # Only advance recipient index if batch was processed successfully
            recipient_index += len(batch)
            login_index += 1

            if recipient_index < len(recipients):
                print(f"[+] Switching accounts...")
                # Reduced account switch delay
                time.sleep(1 if FAST_MODE else 2)

    if logins_updated:
        update_logins_file(logins)

    print("\n" + "="*60)
    print("📋 FINAL SUMMARY")
    print("="*60)
    print(f"📧 Total Recipients: {len(recipients)}")
    print(f"✅ Successfully Sent: {success_count}")
    print(f"❌ Failed to Send: {fail_count}")
    print(f"📈 Success Rate: {(success_count/len(recipients)*100):.1f}%")
    print(f"🚀 Performance Mode: {'FAST' if FAST_MODE else 'NORMAL'}")
    if failed_logins:
        print(f"⚠️ Failed Logins: {', '.join(failed_logins)}")
    print("="*60)

def main():
    """Main execution function"""
    if not ENABLE_MULTI_INSTANCE:
        run_single_instance()
        return
    
    print("="*80)
    print("🚀 MULTI-INSTANCE EMAIL AUTOMATION LAUNCHER")
    print("="*80)
    
    # Create tdata directory for temporary files
    temp_dir = Path("tdata")
    temp_dir.mkdir(exist_ok=True)
    print(f"📁 Created temporary directory: {temp_dir}")
    
    # Clear the shared log file
    with open(LOG_FILE, "w") as f:
        f.write(f"# Multi-Instance Email Automation Log - Started {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    
    # Split data in memory
    instance_data = split_data_in_memory()
    if not instance_data:
        print("❌ Failed to load data")
        return
    
    print(f"\n🚀 Launching {NUM_INSTANCES} instances...")
    
    # Create temporary files for each instance in tdata folder
    for data in instance_data:
        instance_num = data['instance_num']
        
        # Write temp files for this instance in tdata directory
        logins_file = temp_dir / f"logins{instance_num}.txt"
        input_file = temp_dir / f"input{instance_num}.txt"
        
        with open(logins_file, "w") as f:
            f.write("\n".join(data['logins']))
        
        with open(input_file, "w", encoding='utf-8') as f:
            f.write("\n".join(data['recipients']))
        
        print(f"📄 Created temp files for Instance {instance_num}: {logins_file.name}, {input_file.name}")
    
    # Launch instances as separate processes
    processes = []
    for data in instance_data:
        instance_num = data['instance_num']
        
        # Modify environment for this instance - use paths in tdata folder
        env = os.environ.copy()
        env['INSTANCE_NUM'] = str(instance_num)
        env['LOGINS_FILE'] = str(temp_dir / f'logins{instance_num}.txt')
        env['RECIPIENTS_FILE'] = str(temp_dir / f'input{instance_num}.txt')
        
        # Launch instance
        cmd = [sys.executable, __file__, '--instance', str(instance_num)]
        process = subprocess.Popen(cmd, env=env)
        processes.append(process)
        
        print(f"✅ Instance {instance_num} started (PID: {process.pid})")
        
        # Stagger the launches
        if instance_num < NUM_INSTANCES:
            print(f"⏳ Waiting {STAGGER_DELAY}s before next instance...")
            time.sleep(STAGGER_DELAY)
    
    print(f"\n📊 All {NUM_INSTANCES} instances launched! Starting monitor...")
    
    # Start monitoring in a separate thread
    monitor_thread = threading.Thread(target=monitor_instances, daemon=True)
    monitor_thread.start()
    
    # Wait for all processes to complete
    try:
        for process in processes:
            process.wait()
        
        # Clean up temp files after all instances complete
        cleanup_temp_files()
        
        # Wait a bit for monitor to show final results
        time.sleep(2)
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
        
        # Clean up even if interrupted
        cleanup_temp_files()
        print("🛑 All instances stopped and cleaned up.")

def run_instance():
    """Run a single instance when called with --instance flag"""
    instance_num = int(os.environ.get('INSTANCE_NUM', '1'))
    
    # Override global variables for this instance
    global LOGINS_FILE, RECIPIENTS_FILE
    LOGINS_FILE = os.environ.get('LOGINS_FILE', f'tdata/logins{instance_num}.txt')
    RECIPIENTS_FILE = os.environ.get('RECIPIENTS_FILE', f'tdata/input{instance_num}.txt')
    
    print(f"="*60)
    print(f"🤖 EMAIL AUTOMATION INSTANCE {instance_num}")
    print(f"="*60)
    print(f"📂 Using logins: {LOGINS_FILE}")
    print(f"📂 Using recipients: {RECIPIENTS_FILE}")
    print(f"📂 Logging to: {LOG_FILE}")
    print(f"🖥️ Headless mode: {HEADLESS}")
    print(f"⚡ Fast mode: {FAST_MODE}")
    print(f"="*60)
    
    # Run the single instance
    run_single_instance()

if __name__ == "__main__":
    # Check if this is an instance call
    if len(sys.argv) > 2 and sys.argv[1] == '--instance':
        run_instance()
    else:
        main()