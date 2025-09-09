import time
import os
from pathlib import Path
from datetime import datetime
from playwright.sync_api import sync_playwright

# === CONFIGURATION ===
LOGINS_FILE = "logins.txt"
HEADLESS = False  # Set to True for production
LOG_FILE = "deletion_log.txt"

# Subject pattern to match (the dynamic part will be handled with partial matching)
SUBJECT_PATTERN = "Overdue Strategic Planning Services Invoice INV172573"
SUBJECT_PREFIX = "INV - "
SUBJECT_SUFFIX = " Overdue Strategic Planning Services Invoice INV172573"

# Performance settings
TIMEOUT_NORMAL = 10000
TIMEOUT_FAST = 5000
PAGE_LOAD_WAIT = 3000
DELETE_CONFIRMATION_WAIT = 2000

# === TRACKING ===
total_deleted = 0
deletion_log = []

def load_logins():
    """Load login credentials from file"""
    logins = []
    try:
        with open(LOGINS_FILE, 'r') as f:
            for line in f:
                line = line.strip()
                if line and ':' in line:
                    parts = line.split(':', 2)
                    if len(parts) >= 2:
                        email = parts[0].strip()
                        password = parts[1].strip()
                        logins.append((email, password))
    except FileNotFoundError:
        print(f"[!] {LOGINS_FILE} not found!")
        return []
    
    print(f"[+] Loaded {len(logins)} login accounts")
    return logins

def write_log(account, folder, subject, status, error_msg=""):
    """Write deletion results to log file"""
    global deletion_log
    
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    log_entry = f"{timestamp} | Account: {account} | Folder: {folder} | Subject: {subject} | Status: {status}"
    if error_msg:
        log_entry += f" | Error: {error_msg}"
    
    deletion_log.append(log_entry)
    
    # Write to file immediately
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(log_entry + "\n")

def login_to_outlook(email, password, page):
    """Login to Outlook and return success status"""
    try:
        print(f"[>] Logging in to: {email}")
        
        page.goto("https://outlook.office.com/mail", timeout=30000)
        page.wait_for_timeout(PAGE_LOAD_WAIT)
        
        # Handle email input
        email_input = page.locator("input[type='email']")
        if email_input.is_visible():
            email_input.fill(email)
            page.click("input[type='submit']")
            page.wait_for_timeout(2000)
        
        # Handle account selection if needed
        try:
            if page.locator("text=It looks like this email is used with more than one account").is_visible():
                work_account = page.locator("text=Work or school account")
                if work_account.is_visible():
                    work_account.click()
                    page.wait_for_timeout(2000)
        except:
            pass
        
        # Handle password input
        try:
            page.wait_for_selector("input[type='password']", timeout=TIMEOUT_FAST)
            page.fill("input[type='password']", password)
            page.click("input[type='submit']")
            page.wait_for_timeout(3000)
        except Exception as e:
            print(f"[!] Password input failed: {e}")
            return False
        
        # Handle "Stay signed in" prompt
        try:
            if page.locator("input[id='idBtn_Back']").is_visible():
                page.click("input[id='idBtn_Back']")
                page.wait_for_timeout(2000)
        except:
            pass
        
        # Wait for Outlook interface to load
        page.wait_for_timeout(5000)
        
        # Check if we successfully reached the mail interface
        try:
            page.wait_for_selector("button[aria-label='New mail'], [aria-label*='mail'], div[role='main']", timeout=15000)
            print(f"[✓] Successfully logged in to: {email}")
            return True
        except:
            print(f"[!] Failed to reach Outlook interface for: {email}")
            return False
            
    except Exception as e:
        print(f"[!] Login failed for {email}: {e}")
        return False

def find_emails_by_subject(page, folder_name, folder_url):
    """Find emails matching the subject pattern in specified folder"""
    try:
        print(f"[>] Searching for emails in {folder_name}...")
        
        # Navigate to the specified folder
        page.goto(folder_url, wait_until="networkidle", timeout=20000)
        page.wait_for_timeout(PAGE_LOAD_WAIT)
        
        # Wait for email list to load
        email_list_selectors = [
            "[role='listbox'] [role='option']",
            "[data-testid='mail-list'] div[role='listitem']",
            ".ms-List .ms-List-cell",
            "[role='grid'] [role='row']",
            "div[class*='_3gzbu9n']",  # Additional Outlook selector
            "div[data-testid*='message']",
            "[aria-label*='Message list']",
            "div[role='listitem']"
        ]
        
        # Find email container
        email_container = None
        for selector in email_list_selectors:
            try:
                container = page.locator(selector)
                if container.count() > 0:
                    email_container = container
                    print(f"[>] Found {container.count()} emails using selector: {selector}")
                    break
            except:
                continue
        
        if not email_container:
            print(f"[!] No email container found in {folder_name}")
            # Let's try to see what's actually on the page
            try:
                page_text = page.locator("body").text_content()
                if "INV -" in page_text:
                    print(f"[>] Page contains 'INV -' text, trying alternative detection...")
                    # Try to find any element containing our keywords
                    inv_elements = page.locator("text=INV -").all()
                    print(f"[>] Found {len(inv_elements)} elements containing 'INV -'")
                    return []
                else:
                    print(f"[!] No 'INV -' text found on page")
            except:
                pass
            return []
        
        # Search for matching emails
        matching_emails = []
        total_emails = email_container.count()
        
        print(f"[>] Scanning {total_emails} emails for subject pattern...")
        print(f"[>] Looking for patterns: 'INV -', 'Advisory Services', '{SUBJECT_PATTERN}'")
        
        for i in range(total_emails):
            try:
                email_element = email_container.nth(i)
                
                # Get all text content from the email element
                full_text = ""
                try:
                    full_text = email_element.text_content() or ""
                except:
                    pass
                
                # Get aria-label if available
                aria_label = ""
                try:
                    aria_label = email_element.get_attribute("aria-label") or ""
                except:
                    pass
                
                # Get title if available
                title_attr = ""
                try:
                    title_attr = email_element.get_attribute("title") or ""
                except:
                    pass
                
                # Combine all text sources
                combined_text = f"{full_text} {aria_label} {title_attr}".lower()
                
                # Debug: Print first few emails to see what we're getting
                if i < 3:
                    print(f"[DEBUG] Email {i+1} text: {combined_text[:100]}...")
                
                # Check for our patterns (case insensitive)
                patterns_to_check = [
                    "inv -",
                    "advisory services",
                    "pending, please advise",
                    "q1 & q2 2025"
                ]
                
                matches_found = 0
                for pattern in patterns_to_check:
                    if pattern in combined_text:
                        matches_found += 1
                
                # If we find at least 2 patterns, consider it a match
                if matches_found >= 2:
                    # Extract subject for logging
                    subject_text = combined_text[:200]  # First 200 chars for logging
                    
                    matching_emails.append({
                        'element': email_element,
                        'subject': subject_text.strip(),
                        'index': i
                    })
                    print(f"[✓] Found matching email {len(matching_emails)}: {subject_text[:80]}...")
                
                # Alternative check - look specifically for "INV -" pattern
                elif "inv -" in combined_text and ("advisory" in combined_text or "services" in combined_text):
                    subject_text = combined_text[:200]
                    matching_emails.append({
                        'element': email_element,
                        'subject': subject_text.strip(),
                        'index': i
                    })
                    print(f"[✓] Found matching email (alt): {subject_text[:80]}...")
                
            except Exception as e:
                print(f"[!] Error processing email {i}: {e}")
                continue
        
        print(f"[+] Found {len(matching_emails)} matching emails in {folder_name}")
        
        # If we found no matches, let's do one more debug check
        if len(matching_emails) == 0:
            print("[DEBUG] No matches found. Let's check what subjects are actually there...")
            try:
                # Look for any element containing "INV" on the page
                inv_elements = page.locator("text=INV").all()
                print(f"[DEBUG] Found {len(inv_elements)} elements with 'INV' text")
                
                for i, elem in enumerate(inv_elements[:5]):  # Check first 5
                    try:
                        text = elem.text_content()[:100]
                        print(f"[DEBUG] INV element {i+1}: {text}")
                    except:
                        pass
                        
            except Exception as debug_error:
                print(f"[DEBUG] Debug check failed: {debug_error}")
        
        return matching_emails
        
    except Exception as e:
        print(f"[!] Error searching {folder_name}: {e}")
        return []

def delete_email(page, email_info, folder_name):
    """Delete a specific email"""
    try:
        email_element = email_info['element']
        subject = email_info['subject']
        
        print(f"[>] Deleting email: {subject[:50]}...")
        
        # Method 1: Right-click context menu
        try:
            email_element.scroll_into_view_if_needed()
            page.wait_for_timeout(500)
            
            # Right-click to open context menu
            email_element.click(button="right")
            page.wait_for_timeout(1000)
            
            # Look for delete option in context menu
            delete_options = [
                "text=Delete",
                "[aria-label*='Delete']",
                "button:has-text('Delete')",
                "[role='menuitem']:has-text('Delete')"
            ]
            
            for delete_option in delete_options:
                try:
                    if page.locator(delete_option).is_visible():
                        page.click(delete_option)
                        page.wait_for_timeout(DELETE_CONFIRMATION_WAIT)
                        
                        # Handle any confirmation dialogs
                        confirmation_handled = False
                        page.wait_for_timeout(1500)  # Wait for dialog to fully appear
                        
                        # Check for permanent delete confirmation (Deleted Items folder)
                        if page.locator("text=Do you want to permanently delete").is_visible():
                            print("[>] Permanent deletion confirmation found - clicking OK")
                            
                            # Try multiple OK button selectors with the exact structure from the HTML
                            ok_selectors = [
                                "button.fui-Button:has-text('OK')",
                                ".fui-DialogActions button:has-text('OK')",
                                "button[class*='fui-Button']:has-text('OK')",
                                ".fui-DialogBody button:has-text('OK')",
                                "[role='dialog'] .fui-DialogActions button:has-text('OK')",
                                "button[class*='r1alrhcs']:has-text('OK')",
                                "button[type='button']:has-text('OK')",
                                ".rhfpeu0 button:has-text('OK')",
                                "div.fui-DialogActions button:has-text('OK')"
                            ]
                            
                            # First, wait a bit more for dialog to be fully interactive
                            page.wait_for_timeout(2000)
                            
                            for ok_selector in ok_selectors:
                                try:
                                    ok_buttons = page.locator(ok_selector)
                                    button_count = ok_buttons.count()
                                    print(f"[>] Selector '{ok_selector}' found {button_count} buttons")
                                    
                                    if button_count > 0:
                                        ok_button = ok_buttons.first
                                        
                                        # Verify it's actually visible and contains OK
                                        if ok_button.is_visible():
                                            button_text = ok_button.text_content()
                                            print(f"[>] Found visible button with text: '{button_text}'")
                                            
                                            if "OK" in button_text:
                                                print(f"[>] Confirmed OK button found, attempting click...")
                                                
                                                # Method 1: Scroll into view and click
                                                try:
                                                    ok_button.scroll_into_view_if_needed()
                                                    page.wait_for_timeout(1000)
                                                    ok_button.click(timeout=5000)
                                                    page.wait_for_timeout(4000)
                                                    
                                                    if not page.locator("text=Do you want to permanently delete").is_visible():
                                                        print("[✓] OK clicked successfully (scroll+click)")
                                                        confirmation_handled = True
                                                        break
                                                except Exception as e:
                                                    print(f"[!] Scroll+click failed: {e}")
                                                
                                                # Method 2: Force click without waiting for actionability
                                                if not confirmation_handled:
                                                    try:
                                                        ok_button.click(force=True, timeout=5000)
                                                        page.wait_for_timeout(4000)
                                                        if not page.locator("text=Do you want to permanently delete").is_visible():
                                                            print("[✓] OK clicked successfully (force)")
                                                            confirmation_handled = True
                                                            break
                                                    except Exception as e:
                                                        print(f"[!] Force click failed: {e}")
                                                
                                                # Method 3: Direct JavaScript execution
                                                if not confirmation_handled:
                                                    try:
                                                        page.evaluate("""
                                                            () => {
                                                                const buttons = document.querySelectorAll('button.fui-Button');
                                                                for (let btn of buttons) {
                                                                    if (btn.textContent && btn.textContent.includes('OK')) {
                                                                        btn.focus();
                                                                        btn.click();
                                                                        return true;
                                                                    }
                                                                }
                                                                return false;
                                                            }
                                                        """)
                                                        page.wait_for_timeout(4000)
                                                        if not page.locator("text=Do you want to permanently delete").is_visible():
                                                            print("[✓] OK clicked successfully (JavaScript)")
                                                            confirmation_handled = True
                                                            break
                                                    except Exception as e:
                                                        print(f"[!] JavaScript click failed: {e}")
                                                
                                                # Method 4: Mouse click with coordinates
                                                if not confirmation_handled:
                                                    try:
                                                        bbox = ok_button.bounding_box()
                                                        if bbox:
                                                            x = bbox['x'] + bbox['width'] / 2
                                                            y = bbox['y'] + bbox['height'] / 2
                                                            print(f"[>] Trying mouse click at coordinates ({x}, {y})")
                                                            page.mouse.click(x, y)
                                                            page.wait_for_timeout(4000)
                                                            if not page.locator("text=Do you want to permanently delete").is_visible():
                                                                print("[✓] OK clicked successfully (mouse)")
                                                                confirmation_handled = True
                                                                break
                                                    except Exception as e:
                                                        print(f"[!] Mouse click failed: {e}")
                                        
                                except Exception as e:
                                    print(f"[!] OK selector {ok_selector} failed: {e}")
                                    continue
                            
                            # Final fallback - try Enter key
                            if not confirmation_handled:
                                print("[>] All OK button clicks failed, trying Enter key")
                                try:
                                    page.keyboard.press("Enter")
                                    page.wait_for_timeout(2000)
                                    if not page.locator("text=Do you want to permanently delete").is_visible():
                                        print("[✓] Enter key worked")
                                        confirmation_handled = True
                                except:
                                    pass
                        
                        # Check for regular delete confirmation
                        elif page.locator("button:has-text('Delete')").is_visible():
                            dialog_text = ""
                            try:
                                dialog_element = page.locator("[role='dialog']")
                                if dialog_element.is_visible():
                                    dialog_text = dialog_element.text_content().lower()
                            except:
                                pass
                            
                            if "all" not in dialog_text or "permanently delete the selected conversations" in dialog_text:
                                page.click("button:has-text('Delete')")
                                page.wait_for_timeout(1000)
                                print("[✓] Regular deletion confirmed")
                                confirmation_handled = True
                            else:
                                print("[!] Dangerous 'Delete all' dialog - cancelling")
                                page.keyboard.press("Escape")
                                return False
                        
                        if confirmation_handled:
                            print("[✓] Email deleted via context menu")
                            return True
                        else:
                            print("[!] Could not handle confirmation dialog")
                            page.keyboard.press("Escape")  # Cancel if we can't handle it
                            return False
                except:
                    continue
            
        except Exception as e:
            print(f"[!] Context menu method failed: {e}")
        
        # Method 2: Select and use keyboard
        try:
            print("[>] Trying select and keyboard delete...")
            email_element.click()
            page.wait_for_timeout(500)
            
            # Press Delete key
            page.keyboard.press("Delete")
            page.wait_for_timeout(DELETE_CONFIRMATION_WAIT)
            
            # Handle confirmation if needed
            confirmation_handled = False
            page.wait_for_timeout(1500)  # Wait for dialog to fully appear
            
            # Check for permanent delete confirmation (Deleted Items folder)
            if page.locator("text=Do you want to permanently delete").is_visible():
                print("[>] Permanent deletion confirmation found - clicking OK")
                
                # Try multiple OK button selectors with the exact structure from the HTML
                ok_selectors = [
                    "button.fui-Button:has-text('OK')",
                    ".fui-DialogActions button:has-text('OK')",
                    "button[class*='fui-Button']:has-text('OK')",
                    ".fui-DialogBody button:has-text('OK')",
                    "[role='dialog'] .fui-DialogActions button:has-text('OK')",
                    "button[class*='r1alrhcs']:has-text('OK')",
                    "button[type='button']:has-text('OK')",
                    ".rhfpeu0 button:has-text('OK')",
                    "div.fui-DialogActions button:has-text('OK')"
                ]
                
                for ok_selector in ok_selectors:
                    try:
                        ok_buttons = page.locator(ok_selector)
                        button_count = ok_buttons.count()
                        
                        if button_count > 0:
                            ok_button = ok_buttons.first
                            
                            if ok_button.is_visible():
                                button_text = ok_button.text_content()
                                
                                if "OK" in button_text:
                                    print(f"[>] Keyboard method: Found OK button, attempting click...")
                                    
                                    # Try the same 4 methods as context menu
                                    click_methods = [
                                        ("scroll+click", lambda: (
                                            ok_button.scroll_into_view_if_needed(),
                                            page.wait_for_timeout(1000),
                                            ok_button.click(timeout=5000)
                                        )),
                                        ("force", lambda: ok_button.click(force=True, timeout=5000)),
                                        ("javascript", lambda: page.evaluate("""
                                            () => {
                                                const buttons = document.querySelectorAll('button.fui-Button');
                                                for (let btn of buttons) {
                                                    if (btn.textContent && btn.textContent.includes('OK')) {
                                                        btn.focus();
                                                        btn.click();
                                                        return true;
                                                    }
                                                }
                                                return false;
                                            }
                                        """)),
                                        ("mouse", lambda: (
                                            lambda bbox=ok_button.bounding_box(): page.mouse.click(
                                                bbox['x'] + bbox['width'] / 2,
                                                bbox['y'] + bbox['height'] / 2
                                            ) if bbox else None
                                        )())
                                    ]
                                    
                                    for method_name, click_method in click_methods:
                                        try:
                                            click_method()
                                            page.wait_for_timeout(4000)
                                            if not page.locator("text=Do you want to permanently delete").is_visible():
                                                print(f"[✓] OK clicked successfully ({method_name})")
                                                confirmation_handled = True
                                                break
                                        except Exception as e:
                                            print(f"[!] Keyboard method {method_name} failed: {e}")
                                            continue
                                    
                                    if confirmation_handled:
                                        break
                        
                    except Exception as e:
                        continue
                
                # Final fallback - try Enter key
                if not confirmation_handled:
                    print("[>] All OK button clicks failed, trying Enter key")
                    try:
                        page.keyboard.press("Enter")
                        page.wait_for_timeout(2000)
                        if not page.locator("text=Do you want to permanently delete").is_visible():
                            print("[✓] Enter key worked")
                            confirmation_handled = True
                    except:
                        pass
            
            # Check for regular delete confirmation
            elif page.locator("button:has-text('Delete')").is_visible():
                dialog_text = ""
                try:
                    dialog_element = page.locator("[role='dialog']")
                    if dialog_element.is_visible():
                        dialog_text = dialog_element.text_content().lower()
                except:
                    pass
                
                if "all" not in dialog_text or "permanently delete the selected conversations" in dialog_text:
                    page.click("button:has-text('Delete')")
                    page.wait_for_timeout(1000)
                    print("[✓] Regular deletion confirmed")
                    confirmation_handled = True
                else:
                    print("[!] Dangerous 'Delete all' dialog - cancelling")
                    page.keyboard.press("Escape")
                    return False
            
            if confirmation_handled:
                print("[✓] Email deleted via keyboard")
                return True
            else:
                print("[!] Could not handle confirmation dialog in keyboard method")
                page.keyboard.press("Escape")  # Cancel if we can't handle it
                return False
                
        except Exception as e:
            print(f"[!] Keyboard method failed: {e}")
        
        return False
        
    except Exception as e:
        print(f"[!] Error deleting email: {e}")
        return False

def delete_emails_from_folder(page, account, folder_name, folder_url):
    """Delete all matching emails from a specific folder with pagination handling"""
    folder_deleted_count = 0
    max_cycles = 20  # Increased for pagination
    cycle = 0
    
    try:
        while cycle < max_cycles:
            cycle += 1
            print(f"[>] Deletion cycle {cycle} for {folder_name}...")
            
            # Navigate to folder fresh each cycle to handle pagination
            page.goto(folder_url, wait_until="networkidle", timeout=15000)
            page.wait_for_timeout(3000)
            
            # Find matching emails in current view
            matching_emails = find_emails_by_subject(page, folder_name, folder_url)
            
            if not matching_emails:
                if cycle == 1:
                    print(f"[>] No matching emails found in {folder_name}")
                else:
                    print(f"[+] All matching emails cleared from {folder_name}")
                break
            
            print(f"[>] Found {len(matching_emails)} emails to delete in cycle {cycle}")
            cycle_deleted = 0
            
            # Method 1: Individual deletion (more reliable for multiple pages)
            if len(matching_emails) <= 3:
                print("[>] Using individual deletion for small batch...")
                
                for i, email_info in enumerate(matching_emails):
                    try:
                        print(f"[>] Deleting email {i+1}/{len(matching_emails)}: {email_info['subject'][:50]}...")
                        
                        if delete_email(page, email_info, folder_name):
                            cycle_deleted += 1
                            write_log(account, folder_name, email_info['subject'], "DELETED")
                        else:
                            write_log(account, folder_name, email_info['subject'], "FAILED", "Could not delete email")
                        
                        # Small delay between individual deletions
                        page.wait_for_timeout(1000)
                        
                    except Exception as e:
                        error_msg = f"Deletion error: {e}"
                        print(f"[!] {error_msg}")
                        write_log(account, folder_name, email_info['subject'], "FAILED", error_msg)
            
            # Method 2: Small bulk deletion (more conservative)
            else:
                print("[>] Using small bulk deletion...")
                
                # Process in smaller chunks to avoid issues
                chunk_size = min(5, len(matching_emails))  # Reduced from 10 to 5
                emails_to_process = matching_emails[:chunk_size]
                
                try:
                    print(f"[>] Selecting {len(emails_to_process)} emails for bulk deletion...")
                    
                    # Click first email
                    first_email = emails_to_process[0]['element']
                    first_email.scroll_into_view_if_needed()
                    page.wait_for_timeout(500)
                    first_email.click()
                    page.wait_for_timeout(1000)
                    print("[>] Selected first email")
                    
                    # Add others with Ctrl+click
                    for i in range(1, len(emails_to_process)):
                        try:
                            page.keyboard.down("Control")
                            emails_to_process[i]['element'].scroll_into_view_if_needed()
                            page.wait_for_timeout(300)
                            emails_to_process[i]['element'].click()
                            page.wait_for_timeout(300)
                            page.keyboard.up("Control")
                            print(f"[>] Added email {i+1} to selection")
                        except Exception as sel_error:
                            print(f"[!] Selection error for email {i+1}: {sel_error}")
                            page.keyboard.up("Control")  # Make sure Ctrl is released
                            break
                    
                    # Delete selected emails
                    print("[>] Pressing Delete key for selected emails...")
                    page.keyboard.press("Delete")
                    page.wait_for_timeout(2000)
                    
                    # Handle confirmation dialog
                    confirmation_handled = False
                    
                    # Look for the permanent delete confirmation
                    if page.locator("text=Do you want to permanently delete").is_visible():
                        print("[>] Permanent deletion confirmation found - attempting to click OK")
                        
                        # Try clicking OK with more robust method
                        ok_clicked = False
                        
                        # Method 1: Simple approach first
                        try:
                            ok_btn = page.locator("button:has-text('OK'):visible").first
                            if ok_btn.is_visible():
                                print("[>] Trying simple OK click...")
                                ok_btn.click()
                                page.wait_for_timeout(4000)  # Longer wait
                                
                                if not page.locator("text=Do you want to permanently delete").is_visible():
                                    print("[✓] OK clicked successfully")
                                    ok_clicked = True
                                    confirmation_handled = True
                        except:
                            pass
                        
                        # Method 2: Tab navigation + Enter
                        if not ok_clicked:
                            try:
                                print("[>] Trying Tab + Enter method...")
                                page.keyboard.press("Tab")  # Move to OK button
                                page.wait_for_timeout(500)
                                page.keyboard.press("Enter")  # Press it
                                page.wait_for_timeout(4000)
                                
                                if not page.locator("text=Do you want to permanently delete").is_visible():
                                    print("[✓] Tab+Enter worked")
                                    ok_clicked = True
                                    confirmation_handled = True
                            except:
                                pass
                        
                        # Method 3: Space bar (some buttons respond to space)
                        if not ok_clicked:
                            try:
                                print("[>] Trying Space bar method...")
                                page.keyboard.press("Space")
                                page.wait_for_timeout(4000)
                                
                                if not page.locator("text=Do you want to permanently delete").is_visible():
                                    print("[✓] Space bar worked")
                                    ok_clicked = True
                                    confirmation_handled = True
                            except:
                                pass
                        
                        if not ok_clicked:
                            print("[!] Could not click OK button - cancelling")
                            page.keyboard.press("Escape")
                            continue
                    
                    if confirmation_handled:
                        # Log successful bulk deletion
                        for email_info in emails_to_process:
                            write_log(account, folder_name, email_info['subject'], "DELETED (BULK)")
                            cycle_deleted += 1
                        
                        print(f"[+] Bulk deleted {cycle_deleted} emails")
                    
                except Exception as bulk_error:
                    print(f"[!] Bulk deletion failed: {bulk_error}")
                    # Fallback to individual deletion
                    page.keyboard.press("Escape")  # Cancel any open dialogs
                    page.wait_for_timeout(1000)
                    
                    print("[>] Falling back to individual deletion...")
                    for email_info in emails_to_process[:3]:  # Just try first 3
                        try:
                            if delete_email(page, email_info, folder_name):
                                cycle_deleted += 1
                                write_log(account, folder_name, email_info['subject'], "DELETED")
                        except:
                            write_log(account, folder_name, email_info['subject'], "FAILED", "Fallback deletion failed")
            
            folder_deleted_count += cycle_deleted
            print(f"[+] Cycle {cycle}: Deleted {cycle_deleted} emails (Total: {folder_deleted_count})")
            
            # If we didn't delete anything this cycle, break to avoid infinite loop
            if cycle_deleted == 0:
                print(f"[!] No deletions in cycle {cycle}, stopping")
                break
            
            # Wait before next cycle
            if cycle < max_cycles:
                print(f"[>] Waiting before next cycle...")
                page.wait_for_timeout(2000)
        
        print(f"[+] Total deleted from {folder_name}: {folder_deleted_count} emails in {cycle} cycles")
        return folder_deleted_count
        
    except Exception as e:
        print(f"[!] Error processing {folder_name}: {e}")
        return folder_deleted_count

def process_account(email, password, p):
    """Process a single account - delete emails from Sent and Deleted Items"""
    account_deleted_count = 0
    
    try:
        browser = p.chromium.launch(headless=HEADLESS)
        context = browser.new_context()
        page = context.new_page()
        
        # Login
        if not login_to_outlook(email, password, page):
            write_log(email, "LOGIN", "N/A", "FAILED", "Could not login to account")
            browser.close()
            return 0
        
        # Define folders to clean
        folders = [
            ("Sent Items", "https://outlook.office.com/mail/sentitems"),
            ("Deleted Items", "https://outlook.office.com/mail/deleteditems")
        ]
        
        # Process each folder
        for folder_name, folder_url in folders:
            print(f"\n[>] Processing {folder_name} for {email}...")
            deleted_count = delete_emails_from_folder(page, email, folder_name, folder_url)
            account_deleted_count += deleted_count
            
            # Wait between folders
            if deleted_count > 0:
                time.sleep(2)
        
        browser.close()
        return account_deleted_count
        
    except Exception as e:
        print(f"[!] Error processing account {email}: {e}")
        write_log(email, "ACCOUNT", "N/A", "ERROR", str(e))
        try:
            browser.close()
        except:
            pass
        return 0

def main():
    """Main execution function"""
    global total_deleted
    
    print("="*70)
    print("🗑️  EMAIL DELETION SCRIPT")
    print("="*70)
    print(f"Target pattern: INV - [Company] Advisory Services (Q1 & Q2 2025) – Pending, Please Advise")
    print(f"Folders: Sent Items, Deleted Items")
    print(f"Headless mode: {HEADLESS}")
    print("="*70)
    
    # Initialize log file
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write(f"# Email Deletion Log - Started {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"# Target: Advisory Services (Q1 & Q2 2025) emails\n")
        f.write(f"# Folders: Sent Items, Deleted Items\n\n")
    
    # Load accounts
    logins = load_logins()
    if not logins:
        print("[!] No login accounts found. Exiting.")
        return
    
    print(f"[+] Processing {len(logins)} accounts...\n")
    
    # Process each account
    with sync_playwright() as p:
        for i, (email, password) in enumerate(logins, 1):
            print(f"\n{'='*50}")
            print(f"[{i}/{len(logins)}] Processing: {email}")
            print(f"{'='*50}")
            
            account_deleted = process_account(email, password, p)
            total_deleted += account_deleted
            
            print(f"[+] Account {email}: {account_deleted} emails deleted")
            
            # Wait between accounts
            if i < len(logins):
                print(f"[>] Waiting before next account...")
                time.sleep(3)
    
    # Final summary
    print("\n" + "="*70)
    print("📋 FINAL DELETION SUMMARY")
    print("="*70)
    print(f"🏦 Accounts Processed: {len(logins)}")
    print(f"🗑️  Total Emails Deleted: {total_deleted}")
    print(f"📄 Detailed log: {LOG_FILE}")
    print("="*70)
    
    # Display deletion log summary
    if deletion_log:
        print("\n📝 Recent Deletions:")
        for log_entry in deletion_log[-10:]:  # Show last 10 entries
            print(f"   {log_entry}")
        
        if len(deletion_log) > 10:
            print(f"   ... and {len(deletion_log) - 10} more entries (see {LOG_FILE})")

if __name__ == "__main__":
    main()