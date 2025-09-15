import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import requests
import json
from msal import ConfidentialClientApplication
import threading
import time
from tkinterweb import HtmlFrame
import re
import pickle
import os
from dotenv import load_dotenv
import os

# Load variables from .env file
load_dotenv()

def apply_bold_syntax(message):
    # Replace **text** with span instead of strong for better font control
    return re.sub(r"\*\*(.*?)\*\*", r'<span style="font-family: Calibri, Arial, sans-serif; font-size: 12px; font-weight: bold;">\1</span>', message)


# --- MS Graph API Authentication ---
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
SCOPES = ["https://graph.microsoft.com/.default"]

cached_token = None
token_expiry = 0

use_signature = True

df = None
is_sending = False  # Flag to track email sending status

# Threading support - store message threads for follow-up emails
message_threads = {}  # Store message IDs for threading {recipient_email: message_id}
conversation_ids = {}  # Store conversation IDs {recipient_email: conversation_id}
thread_file = "email_threads.pkl"  # File to persist thread data

# Sample emails array - you can modify these emails as needed
available_emails = [
    ("kathiravan.m@codework.ai", "kathiravan.m@codework.ai"),
    ("jagadeesh.rp@codework.ai", "jagadeesh.rp@codework.ai"),
    ("james@codeworkx.ai", "james@codeworkx.ai"),
    ("letty@thecodeworx.com", "letty@thecodeworx.com"),
    ("justin@procodework.in", "justin@procodework.in"),
    ("Jagadeesh@codework.life", "Jagadeesh@codework.life"),
]

# Sender signature mapping - maps email addresses to signature information
sender_signatures = {
    "jagadeesh.rp@codework.ai": {
        "name": "Jagadeesh",
        "title": "Business Development Manager", 
        "email": "jagadeesh.rp@codework.ai"
    },
    "james@codeworkx.ai": {
        "name": "James",
        "title": "Business Development Manager",
        "email": "james@codeworkx.ai"
    },
    "letty@thecodeworx.com": {
        "name": "Letty",
        "title": "Business Development Manager",
        "email": "letty@thecodeworx.com"
    },
    "justin@procodework.in": {
        "name": "Justin",
        "title": "Business Development Manager",
        "email": "justin@procodework.in"
    },
    "Jagadeesh@codework.life": {
        "name": "Jagadeesh",
        "title": "Business Development Manager",
        "email": "Jagadeesh@codework.life"
    }
}

# Time intervals array - you can modify these intervals as needed
time_intervals = [
    ("15 seconds", 15),
    ("30 seconds", 30),
    ("1 minute", 60),
    ("2 minutes", 120),
    ("3 minutes", 180),
    ("5 minutes", 300)
]

# --- Threading Helper Functions ---
def save_thread_data():
    """Save message threads to file for persistence"""
    try:
        with open(thread_file, 'wb') as f:
            pickle.dump({'threads': message_threads, 'conversations': conversation_ids}, f)
    except Exception as e:
        print(f"Error saving thread data: {e}")

def load_thread_data():
    """Load message threads from file"""
    global message_threads, conversation_ids
    try:
        if os.path.exists(thread_file):
            with open(thread_file, 'rb') as f:
                data = pickle.load(f)
                message_threads = data.get('threads', {})
                conversation_ids = data.get('conversations', {})
                update_status(f"Loaded {len(message_threads)} existing email threads", "blue")
    except Exception as e:
        print(f"Error loading thread data: {e}")
        message_threads = {}
        conversation_ids = {}

def clear_thread_data():
    """Clear all stored thread data"""
    global message_threads, conversation_ids
    message_threads = {}
    conversation_ids = {}
    if os.path.exists(thread_file):
        os.remove(thread_file)
    update_status("Thread data cleared", "orange")

def get_sent_message_details(sender_email, recipient_email, headers):
    """Get details of the most recent sent message with retry logic"""
    max_attempts = 3
    wait_time = 3  # Start with 3 seconds
    
    for attempt in range(max_attempts):
        try:
            # Get the most recent sent message
            url = (f"https://graph.microsoft.com/v1.0/users/{sender_email}/mailFolders/SentItems/messages"
                   f"?$top=1&$orderby=sentDateTime desc"
                   f"&$select=internetMessageId,conversationId,id")

            sent_messages_response = requests.get(url, headers=headers)
            
            if sent_messages_response.status_code == 200:
                sent_data = sent_messages_response.json()
                messages = sent_data.get('value', [])
                
                if messages:
                    message_info = messages[0]
                    update_status(f"Thread info captured for {recipient_email}", "green")
                    return {
                        'internetMessageId': message_info.get('internetMessageId'),
                        'conversationId': message_info.get('conversationId'),
                        'messageId': message_info.get('id')
                    }
            
            # If no message found and not the last attempt, wait and retry
            if attempt < max_attempts - 1:
                update_status(f"Waiting for message to appear in Sent Items... (attempt {attempt + 1})", "orange")
                time.sleep(wait_time)
                wait_time += 2  # Increase wait time for next attempt
                
        except Exception as e:
            print(f"Error getting sent message details (attempt {attempt + 1}): {e}")
            if attempt < max_attempts - 1:
                time.sleep(wait_time)
    
    # If all attempts failed
    update_status(f"Could not capture thread info for {recipient_email}", "red")
    return None

# Also update the email sending section in send_emails() function
# Replace the existing thread capture code with this improved version:

def improved_thread_capture_in_send_emails():
    """
    This shows how to modify the thread capture section in your send_emails() function.
    Replace the existing thread capture code (around lines 400-410) with this:
    """
    
    # After successful email send (status 202), replace the existing code with:
    if response.status_code == 202:
        tree.set(index, "Status", "Sent")
        success_count += 1
        email_type = "follow-up" if is_follow_up_email else "new"
        update_status(f"✓ {email_type.title()} email sent to: {recipient_email}", "green")
        
        # Store message ID for future threading (only for new emails, not follow-ups)
        if not is_follow_up_email:
            update_status(f"Capturing thread info for {recipient_email}...", "blue")
            
            # Get the sent message details with improved retry logic
            message_details = get_sent_message_details(sender_email, recipient_email, headers)
            if message_details and message_details.get('internetMessageId'):
                message_threads[recipient_email] = message_details['internetMessageId']
                conversation_ids[recipient_email] = message_details['conversationId']
                save_thread_data()  # Save to file for persistence
                update_status(f"✓ Thread captured for {recipient_email}", "green")
            else:
                update_status(f"⚠ Thread capture failed for {recipient_email}", "orange")

# --- Helper functions ---
def get_access_token(force_refresh=False):
    global cached_token, token_expiry

    if not force_refresh and cached_token and time.time() < token_expiry - 300:
        return cached_token
    
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=SCOPES)
    if "access_token" in result:
        cached_token = result['access_token']
        token_expiry = time.time() + result.get('expires_in', 3600)
        return cached_token
    else:
        messagebox.showerror("Error", f"Failed to get token: {result.get('error_description')}")
        return None

def toggle_signature(enabled):
    """Toggle signature inclusion on/off"""
    global use_signature
    use_signature = enabled
    status_text = "enabled" if enabled else "disabled"
    update_status(f"Email signature {status_text}", "blue")
    
    # Refresh preview if an email is selected
    if tree.selection():
        preview_email()

def toggle_follow_up_mode(enabled):
    """Toggle follow-up mode (send as replies)"""
    global follow_up_mode
    follow_up_mode = enabled
    mode_text = "Follow-up (Reply)" if enabled else "New Email"
    update_status(f"Email mode: {mode_text}", "blue")
    
    # Update button text based on mode
    if enabled:
        send_button.config(text="Send Follow-up Emails")
        # Show follow-up template section
        followup_frame.pack(fill="x", pady=5, after=message_frame)
    else:
        send_button.config(text="Send New Emails")
        # Hide follow-up template section
        followup_frame.pack_forget()
    
    # Refresh preview
    if tree.selection():
        preview_email()

def populate_email_dropdown():
    """Populate the dropdown with sample emails from the array"""
    global available_emails
   
    # Update the combobox with sample emails
    email_sender_combo['values'] = [entry[0] for entry in available_emails]
    if available_emails:
        email_sender_combo.set(available_emails[0][0])  # Set first email as default
   
    update_status(f"Loaded {len(available_emails)} sample emails", "green")

def populate_time_dropdown():
    """Populate the time interval dropdown with predefined intervals"""
    time_interval_combo['values'] = [interval[0] for interval in time_intervals]
    if time_intervals:
        time_interval_combo.set(time_intervals[2][0])  # Default to "1 minute"

def get_selected_time_interval():
    """Get the selected time interval in seconds"""
    selected_display = time_interval_combo.get()
    for display, seconds in time_intervals:
        if display == selected_display:
            return seconds
    return 60  # Default to 60 seconds if nothing selected

def get_selected_sender_email():
    """Get the actual email address from the selected combo box entry"""
    selected_display = email_sender_combo.get()
    for display, email in available_emails:
        if display == selected_display:
            return email
    return None

def safe_format(row, template):
    try:
        return template.format(**row)
    except KeyError as e:
        return f"Error: Missing field {e} in data"

def format_message_to_html(message):
    """
    Convert plain text message to HTML with table-based formatting for email compatibility
    using the selected font and size.
    """
    selected_font = font_var.get() if 'font_var' in globals() else "Calibri"
    selected_size = size_var.get() if 'size_var' in globals() else "22"

    # Replace double line breaks with table row breaks
    message = message.replace('\n\n',
        f'</td></tr><tr><td style="font-family: {selected_font}, Arial, sans-serif; font-size: {selected_size}px; padding: 0 0 12px 0;">'
    )

    # Replace single line breaks with <br>
    message = message.replace('\n', '<br>')

    # Wrap in table structure
    message = (
        f'<table width="100%" cellpadding="0" cellspacing="0" '
        f'style="font-family: {selected_font}, Arial, sans-serif; font-size: {selected_size}px;">'
        f'<tr><td style="font-family: {selected_font}, Arial, sans-serif; font-size: {selected_size}px; '
        f'padding: 0 0 12px 0;">{message}</td></tr></table>'
    )

    return message

def update_status(message, color="black"):
    """Update status label with message and color"""
    status_label.config(text=message, fg=color)
    root.update()

def update_progress(current, total):
    """Update progress bar"""
    progress_var.set((current / total) * 100)
    progress_label.config(text=f"Progress: {current}/{total} emails sent")
    root.update()

def update_countdown(seconds):
    """Update countdown timer"""
    countdown_label.config(text=f"Next email in: {seconds}s", fg="orange")
    root.update()

def clear_countdown():
    """Clear countdown timer"""
    countdown_label.config(text="", fg="black")
    root.update()

# --- GUI functions ---
def browse_file():
    global df
    filename = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx"), ("Excel files", "*.xls")]
    )
    if not filename:
        return
    entry_excel.delete(0, tk.END)
    entry_excel.insert(0, filename)
   
    try:
        df = pd.read_excel(filename)
        update_status(f"Excel file loaded: {len(df)} rows", "green")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read Excel: {e}")
        update_status("Error loading Excel file", "red")
        return

    if 'email' not in df.columns and 'Email' not in df.columns:
        messagebox.showerror("Error", "Excel must have an 'email' column")
        update_status("Error: No 'email' column found", "red")
        return

    # Setup Treeview dynamically
    tree.delete(*tree.get_children())
   
    # Add Status column to the dataframe columns
    columns_list = list(df.columns) + ["Status"]
    tree["columns"] = columns_list
    tree["show"] = "headings"
   
    for col in columns_list:
        tree.heading(col, text=col)
        if col == "Status":
            tree.column(col, width=80)
        else:
            tree.column(col, width=100)

    # Insert rows with empty status initially
    for index, row in df.iterrows():
        values = [row[col] for col in df.columns] + ["Pending"]
        tree.insert("", "end", iid=index, values=values)

def send_emails_thread():
    global is_sending
    if is_sending:
        messagebox.showwarning("Warning", "Email sending already in progress!")
        return
    threading.Thread(target=send_emails, daemon=True).start()

def send_emails():
    global is_sending, message_threads, conversation_ids
    if df is None or df.empty:
        messagebox.showwarning("Warning", "No Excel data loaded.")
        return
   
    sender_email = get_selected_sender_email()
    wait_time = get_selected_time_interval()
    is_follow_up_mode = follow_up_var.get()  # Check if follow-up mode is enabled
    
    # Get templates based on mode
    if is_follow_up_mode:
        subject_template = entry_followup_subject.get()
        message_template = text_followup_message.get("1.0", tk.END)
        if not subject_template or not message_template.strip():
            messagebox.showwarning("Warning", "Please fill follow-up email subject and message.")
            return
    else:
        subject_template = entry_subject.get()
        message_template = text_message.get("1.0", tk.END)
        if not subject_template or not message_template.strip():
            messagebox.showwarning("Warning", "Please fill initial email subject and message.")
            return
    
    message_template = apply_bold_syntax(message_template)
   
    if not sender_email:
        messagebox.showwarning("Warning", "Please select a sender email address.")
        return

    # For follow-up mode, check if we have threads
    if is_follow_up_mode and not message_threads:
        messagebox.showwarning("Warning", "No email threads available for follow-up. Please send initial emails first.")
        return

    is_sending = True
    send_button.config(state="disabled", text="Sending...", bg="gray")
    update_status("Getting authentication token...", "blue")
   
    token = get_access_token(force_refresh=True)
    if not token:
        is_sending = False
        send_button.config(state="normal", text="Send Emails", bg="green")
        update_status("Authentication failed", "red")
        return

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
   
    success_count = 0
    failed_count = 0
    total = len(df)
   
    progress_var.set(0)
    update_status("Sending emails...", "blue")
   
    for index, row in df.iterrows():
        if not is_sending:
            break
           
        email_col = 'email' if 'email' in df.columns else 'Email'
        recipient_email = row[email_col]
       
        if pd.isna(recipient_email) or not recipient_email:
            tree.set(index, "Status", "No Email")
            failed_count += 1
            continue

        # For follow-up mode, skip if no thread exists for this recipient
        if is_follow_up_mode and recipient_email not in message_threads:
            tree.set(index, "Status", "No Thread")
            failed_count += 1
            continue
           
        update_status(f"Sending {'follow-up' if is_follow_up_mode else 'initial'} email to: {recipient_email}", "blue")
        tree.set(index, "Status", "Sending...")
        root.update()
       
        # Format both subject and message with row data
        try:
            formatted_subject = safe_format(row, subject_template)
            formatted_message = safe_format(row, message_template)
        except Exception as e:
            tree.set(index, "Status", "Format Error")
            failed_count += 1
            continue

        signature_html = get_dynamic_signature(sender_email) if use_signature else ""
        selected_font = font_var.get()
        selected_size = size_var.get()
           
        formatted_message = format_message_to_html(formatted_message)
        formatted_message = f"""
<div style="font-family: '{selected_font}', sans-serif; font-size: {selected_size}px; line-height: 1.4; color: #000000;">
{formatted_message}
{signature_html}
</div>
"""
        
        # Check if this should be sent as a follow-up (reply)
        is_follow_up_email = is_follow_up_mode and recipient_email in message_threads
        
        try:
            if is_follow_up_email:
                # --- Follow-up Email Logic using createReply ---
                original_message_id = message_threads.get(recipient_email)
                if not original_message_id:
                    update_status(f"✗ No message ID for {recipient_email}", "red")
                    failed_count += 1
                    continue

                # 1. Create a reply draft
                create_reply_url = f"https://graph.microsoft.com/v1.0/users/{sender_email}/messages/{original_message_id}/createReply"
                reply_draft_response = requests.post(create_reply_url, headers=headers)

                if reply_draft_response.status_code == 401:
                    token = get_access_token(force_refresh=True)
                    if token:
                        headers["Authorization"] = f"Bearer {token}"
                        reply_draft_response = requests.post(create_reply_url, headers=headers)

                if reply_draft_response.status_code != 201:
                    update_status(f"✗ Failed to create draft for {recipient_email} ({reply_draft_response.status_code})", "red")
                    failed_count += 1
                    continue

                draft_message = reply_draft_response.json()
                draft_id = draft_message.get('id')

                # 2. Update the draft with the new subject and body
                final_subject = formatted_subject
                if not final_subject.startswith("RE:"):
                    final_subject = f"RE: {final_subject}"

                update_draft_url = f"https://graph.microsoft.com/v1.0/users/{sender_email}/messages/{draft_id}"
                draft_update_payload = {
                    "subject": final_subject,
                    "body": {
                        "contentType": "HTML",
                        "content": formatted_message
                    }
                }
                update_response = requests.patch(update_draft_url, headers=headers, data=json.dumps(draft_update_payload))

                if update_response.status_code == 401:
                    token = get_access_token(force_refresh=True)
                    if token:
                        headers["Authorization"] = f"Bearer {token}"
                        update_response = requests.patch(update_draft_url, headers=headers, data=json.dumps(draft_update_payload))

                if update_response.status_code != 200:
                    update_status(f"✗ Failed to update draft for {recipient_email}", "red")
                    failed_count += 1
                    continue

                # 3. Send the draft
                send_draft_url = f"https://graph.microsoft.com/v1.0/users/{sender_email}/messages/{draft_id}/send"
                response = requests.post(send_draft_url, headers=headers)

            else:
                # --- Initial Email Logic ---
                email_data = {
                    "message": {
                        "subject": formatted_subject,
                        "body": {"contentType": "HTML", "content": formatted_message},
                        "toRecipients": [{"emailAddress": {"address": recipient_email}}]
                    }
                }
                response = requests.post(
                    f"https://graph.microsoft.com/v1.0/users/{sender_email}/sendMail",
                    headers=headers,
                    data=json.dumps(email_data)
                )

            if response.status_code == 202:
                tree.set(index, "Status", "Sent")
                success_count += 1
                email_type = "follow-up" if is_follow_up_email else "new"
                update_status(f"✓ {email_type.title()} email sent to: {recipient_email}", "green")
                
                if not is_follow_up_email:
                    time.sleep(3)  # Wait a bit longer for the message to appear in Sent Items
                    message_details = get_sent_message_details(sender_email, recipient_email, headers)
                    if message_details and message_details.get('messageId'):
                        message_threads[recipient_email] = message_details['messageId']
                        conversation_ids[recipient_email] = message_details.get('conversationId')
                        save_thread_data()
                        update_status(f"✓ Thread captured for {recipient_email}", "green")
                    else:
                        update_status(f"⚠ Thread capture failed for {recipient_email}", "orange")

            elif response.status_code == 401:
                # Token expired - try refreshing once
                update_status("Token expired, refreshing...", "orange")
                token = get_access_token(force_refresh=True)
                if token:
                    headers["Authorization"] = f"Bearer {token}"
                    # Retry the request
                    if is_follow_up_email:
                        send_draft_url = f"https://graph.microsoft.com/v1.0/users/{sender_email}/messages/{draft_id}/send"
                        response = requests.post(send_draft_url, headers=headers)
                    else:
                        response = requests.post(
                            f"https://graph.microsoft.com/v1.0/users/{sender_email}/sendMail",
                            headers=headers,
                            data=json.dumps(email_data)
                        )
                    if response.status_code == 202:
                        tree.set(index, "Status", "Sent")
                        success_count += 1
                        update_status(f"✓ Email sent successfully to: {recipient_email} (after token refresh)", "green")
                    else:
                        tree.set(index, "Status", "Failed")
                        failed_count += 1
                        update_status(f"✗ Failed after token refresh: {recipient_email} - {response.status_code}", "red")
                else:
                    tree.set(index, "Status", "Auth Failed")
                    failed_count += 1
                    update_status(f"✗ Token refresh failed for: {recipient_email}", "red")
            
            elif response.status_code == 429:
                # Rate limited - wait and retry
                retry_after = int(response.headers.get('Retry-After', 60))
                update_status(f"Rate limited, waiting {retry_after}s before retry...", "orange")
                for countdown in range(retry_after, 0, -1):
                    if not is_sending:
                        break
                    update_countdown(countdown)
                    time.sleep(1)
                clear_countdown()
                
                # Retry after waiting
                if is_sending:
                    response = requests.post(
                        f"https://graph.microsoft.com/v1.0/users/{sender_email}/sendMail",
                        headers=headers,
                        data=json.dumps(email_data)
                    )
                    if response.status_code == 202:
                        tree.set(index, "Status", "Sent")
                        success_count += 1
                        update_status(f"✓ Email sent successfully to: {recipient_email} (after rate limit)", "green")
                    else:
                        tree.set(index, "Status", "Failed")
                        failed_count += 1
                        update_status(f"✗ Failed after rate limit retry: {recipient_email} - {response.status_code}", "red")
            
            else:
                tree.set(index, "Status", "Failed")
                failed_count += 1
                error_msg = ""
                try:
                    error_response = response.json()
                    error_msg = error_response.get('error', {}).get('message', '')
                except:
                    pass
                update_status(f"✗ Failed to send to: {recipient_email} - {response.status_code} {error_msg}", "red")

        except Exception as e:
            tree.set(index, "Status", "Error")
            failed_count += 1
            update_status(f"✗ Error sending to: {recipient_email} - {str(e)}", "red")
            
        update_progress(success_count + failed_count, total)
       
        # Wait for selected time interval before next email (with countdown)
        if index < len(df) - 1 and is_sending:
            wait_time = get_selected_time_interval()
            for countdown in range(wait_time, 0, -1):
                if not is_sending:
                    break
                update_countdown(countdown)
                time.sleep(1)
            clear_countdown()
   
    is_sending = False
    send_button.config(state="normal", 
                      text="Send Follow-up Emails" if follow_up_var.get() else "Send New Emails", 
                      bg="green")
   
    final_message = f"Email sending completed! ✓ Success: {success_count}, ✗ Failed: {failed_count}"
    update_status(final_message, "green" if failed_count == 0 else "orange")
    messagebox.showinfo("Email Send Complete", f"Emails sent: {success_count}/{total}\nFailed: {failed_count}")

def stop_sending():
    global is_sending
    is_sending = False
    send_button.config(state="normal", text="Send Emails", bg="green")
    update_status("Email sending stopped by user", "orange")

def show_preview(event):
    preview_email()

def preview_email():
    selection = tree.selection()
    if not selection or df is None:
        html_preview.load_html("<p>No email selected or no data loaded.</p>")
        return
    sender_email = get_selected_sender_email()
    try:
        row_id = selection[0]
        row = df.iloc[int(row_id)]
        
        is_follow_up_mode = follow_up_var.get()
        
        # Get the appropriate templates based on mode
        if is_follow_up_mode:
            subject_template = entry_followup_subject.get()
            message_template = text_followup_message.get("1.0", tk.END)
        else:
            subject_template = entry_subject.get()
            message_template = text_message.get("1.0", tk.END)
        
        message_template = apply_bold_syntax(message_template)
       
        if not message_template.strip():
            html_preview.load_html("<p>No message template provided.</p>")
            return
           
        formatted_subject = safe_format(row, subject_template)
        formatted_message = safe_format(row, message_template)
        formatted_message = format_message_to_html(formatted_message)
       
        email_col = 'email' if 'email' in df.columns else 'Email'
        recipient_email = row[email_col] if email_col in row else "No email found"

        signature_html = get_dynamic_signature(sender_email) if use_signature else ""
        selected_font = font_var.get()
        selected_size = size_var.get()
        
        # Check if this would be a follow-up email
        is_follow_up_email = is_follow_up_mode and recipient_email in message_threads
        
        final_subject = formatted_subject
        if is_follow_up_email:
            if not formatted_subject.startswith("RE:"):
                final_subject = f"RE: {formatted_subject}"
        
        # Determine email type for display
        if is_follow_up_mode:
            if is_follow_up_email:
                email_type_indicator = " (Follow-up Reply)"
            else:
                email_type_indicator = " (No Thread Available)"
        else:
            email_type_indicator = " (New Email)"
       
        preview_html = f"""
        <div style="font-family: {selected_font} !important; font-size:{selected_size}px; padding: 10px; border: 1px solid #ccc; margin: 5px;">
            <h3 style="color: #000000; margin-top: 0; font-family: {selected_font} !important; font-size:{selected_size}px;">Email Preview{email_type_indicator}</h3>
            <p style="font-family: {selected_font} !important; font-size:{selected_size}px;"><strong>To:</strong> {recipient_email}</p>
            <p style="font-family: {selected_font} !important; font-size:{selected_size}px;"><strong>Subject:</strong> {final_subject}</p>
            <hr>
            <div style="background: #f9f9f9; padding: 10px; border-left: 3px solid #007acc; font-family: {selected_font} !important; font-size:{selected_size}px;">
                {formatted_message}
                {signature_html}
            </div>
        </div>
        """
        html_preview.load_html(preview_html)
    except Exception as e:
        html_preview.load_html(f"<p>Error generating preview: {str(e)}</p>")

def get_dynamic_signature(sender_email):
    """Generate dynamic signature HTML based on sender email with selected font and size"""
    sender_info = sender_signatures.get(sender_email, {
        "name": "Team Member",
        "title": "Staff",
        "email": sender_email
    })

    selected_font = font_var.get() if 'font_var' in globals() else "Calibri"
    selected_size = size_var.get() if 'size_var' in globals() else "22"

    signature_html = f"""
<table width="100%" cellpadding="0" cellspacing="0" border="0" 
       style="font-family: {selected_font}, Arial, sans-serif; font-size: {selected_size}px; margin-top: 16px;">
<tr>
<td style="font-family: {selected_font}, Arial, sans-serif; font-size: {selected_size}px; line-height: 1.4; color: #000000;">
<p style="font-family: {selected_font}, Arial, sans-serif; font-size: {selected_size}px; margin: 8px 0;">
Thanks,<br>
{sender_info['name']}<br>
<span style="font-family: {selected_font}, Arial, sans-serif; font-size: {selected_size}px; font-weight: bold;">
{sender_info['title']}
</span></p>

<p style="font-family: {selected_font}, Arial, sans-serif; font-size: {selected_size}px; margin: 8px 0;">
<img src="https://codework.ai/Logopng.png" alt="Company Logo" style="height:40px; margin-bottom: 8px;"><br>
<a href="mailto:{sender_info['email']}" 
   style="font-family: {selected_font}, Arial, sans-serif; font-size: {selected_size}px; color: #000237; text-decoration: underline; font-weight: bold;">
   {sender_info['email']}
</a><br>
<a href="https://www.codework.ai" 
   style="font-family: {selected_font}, Arial, sans-serif; font-size: {selected_size}px; color: #000237; text-decoration: underline;">
   www.codework.ai
</a>
</p>

<p style="font-family: {selected_font}, Arial, sans-serif; font-size: {selected_size}px; color: #000237; margin: 8px 0;">
Please click here to 
<a href="mailto:{sender_info['email']}?subject=Please%20unsubscribe" 
   style="font-family: {selected_font}, Arial, sans-serif; font-size: {selected_size}px; color: #0078d4; text-decoration: underline;">
   Unsubscribe
</a>
</p>
</td>
</tr>
</table>
"""
    return signature_html

def show_thread_info():
    """Show information about stored email threads"""
    if message_threads:
        thread_info = f"Stored Email Threads: {len(message_threads)}\n\n"
        for email, msg_id in list(message_threads.items())[:10]:  # Show first 10
            thread_info += f"• {email}\n"
        if len(message_threads) > 10:
            thread_info += f"... and {len(message_threads) - 10} more"
        messagebox.showinfo("Email Thread Information", thread_info)
    else:
        messagebox.showinfo("Email Thread Information", "No email threads stored yet.\n\nThreads are created when you send new emails and can be used for follow-up replies.")

# --- GUI Setup ---
root = tk.Tk()
root.title("Bulk Email Sender with Threading Support")
root.geometry("1400x900")

# Load thread data on startup
load_thread_data()

# Main frame
main_frame = tk.Frame(root)
main_frame.pack(fill="both", expand=True, padx=10, pady=10)

# Left frame for controls
frame_left = tk.Frame(main_frame, width=700)
frame_left.pack(side="left", fill="both", expand=True, padx=(0, 10))
frame_left.pack_propagate(False)

# Right frame for preview
frame_right = tk.Frame(main_frame, width=700)
frame_right.pack(side="right", fill="both", expand=True)

# Top section - File input and email settings
frame_top = tk.Frame(frame_left)
frame_top.pack(fill="x", pady=5)

excel_frame = tk.Frame(frame_top)
excel_frame.pack(fill="x", pady=2)

tk.Label(excel_frame, text="Excel File:").pack(side="left", pady=2)
entry_excel = tk.Entry(excel_frame, width=45)
entry_excel.pack(side="left", padx=(10, 5), pady=2)
tk.Button(excel_frame, text="Browse", command=browse_file).pack(side="left", pady=2)

# Email sender selection
sender_frame = tk.Frame(frame_top)
sender_frame.pack(fill="x", pady=2)

tk.Label(sender_frame, text="Email Sender:").pack(side="left", pady=2)
email_sender_combo = ttk.Combobox(sender_frame, width=45, state="readonly")
email_sender_combo.pack(side="left", padx=(10, 0), pady=2)

# Time interval selection
interval_frame = tk.Frame(frame_top)
interval_frame.pack(fill="x", pady=2)

tk.Label(interval_frame, text="Time Interval Between Emails:").pack(side="left", pady=2)
time_interval_combo = ttk.Combobox(interval_frame, width=25, state="readonly")
time_interval_combo.pack(side="left", padx=(10, 0), pady=2)

# Email mode selection (New vs Follow-up)
mode_frame = tk.Frame(frame_top)
mode_frame.pack(fill="x", pady=2)

tk.Label(mode_frame, text="Email Mode:").pack(side="left", pady=2)
follow_up_var = tk.BooleanVar(value=False)
follow_up_checkbox = tk.Checkbutton(mode_frame, text="Send as Follow-up (Reply to existing threads)", 
                                   variable=follow_up_var, 
                                   command=lambda: toggle_follow_up_mode(follow_up_var.get()))
follow_up_checkbox.pack(side="left", padx=(10, 0), pady=2)

# Threading management buttons
thread_frame = tk.Frame(frame_top)
thread_frame.pack(fill="x", pady=2)

tk.Button(thread_frame, text="Show Thread Info", command=show_thread_info, 
         bg="lightblue", font=("Arial", 8)).pack(side="left", padx=5)
tk.Button(thread_frame, text="Clear Thread Data", command=clear_thread_data, 
         bg="orange", font=("Arial", 8)).pack(side="left", padx=5)

# Signature toggle
signature_frame = tk.Frame(frame_top)
signature_frame.pack(fill="x", pady=2)

tk.Label(signature_frame, text="Include Signature:").pack(side="left", pady=2)
signature_var = tk.BooleanVar(value=True)
signature_checkbox = tk.Checkbutton(signature_frame, text="Enable email signature", 
                                   variable=signature_var, 
                                   command=lambda: toggle_signature(signature_var.get()))
signature_checkbox.pack(side="left", padx=(10, 0), pady=2)

# Font Theme & Size Selection
font_frame = tk.Frame(frame_top)
font_frame.pack(fill="x", pady=2)

tk.Label(font_frame, text="Theme Font:").pack(side="left", pady=2)
font_options = ["Calibri", "Arial", "Times New Roman", "Verdana", "Georgia"]
font_var = tk.StringVar(value="Calibri")
font_combo = ttk.Combobox(font_frame, textvariable=font_var, values=font_options, width=20, state="readonly")
font_combo.pack(side="left", padx=(10, 0), pady=2)
font_combo.bind('<<ComboboxSelected>>', lambda e: preview_email() if tree.selection() else None)

tk.Label(font_frame, text="Font Size:").pack(side="left", padx=(20, 0), pady=2)
size_options = [str(i) for i in range(12, 31)]
size_var = tk.StringVar(value="22")
size_combo = ttk.Combobox(font_frame, textvariable=size_var, values=size_options, width=5, state="readonly")
size_combo.pack(side="left", padx=(10, 0), pady=2)
size_combo.bind('<<ComboboxSelected>>', lambda e: preview_email() if tree.selection() else None)

# Initial Email Template Section
tk.Label(frame_top, text="━━━ Initial Email Template ━━━", font=("Arial", 10, "bold")).pack(fill="x", pady=(10, 2))

subject_frame = tk.Frame(frame_top)
subject_frame.pack(fill="x", pady=2)

tk.Label(subject_frame, text="Initial Subject Template:").pack(side="left", pady=2)
entry_subject = tk.Entry(subject_frame, width=40)
entry_subject.pack(side="left", padx=(10, 0), pady=2)
entry_subject.bind('<KeyRelease>', lambda e: preview_email() if tree.selection() else None)

message_frame = tk.Frame(frame_top)
message_frame.pack(fill="x", pady=2)

tk.Label(message_frame, text="Initial Message Template:", anchor="nw").pack(side="left", pady=2)
text_message = tk.Text(message_frame, width=45, height=6)
text_message.pack(side="left", padx=(10, 0), pady=2)
text_message.bind('<KeyRelease>', lambda e: root.after_idle(lambda: preview_email() if tree.selection() else None))

# Follow-up Email Template Section (initially hidden)
followup_frame = tk.Frame(frame_top)

tk.Label(followup_frame, text="━━━ Follow-up Email Template ━━━", font=("Arial", 10, "bold")).pack(fill="x", pady=(10, 2))

followup_subject_frame = tk.Frame(followup_frame)
followup_subject_frame.pack(fill="x", pady=2)

tk.Label(followup_subject_frame, text="Follow-up Subject Template:").pack(side="left", pady=2)
entry_followup_subject = tk.Entry(followup_subject_frame, width=40)
entry_followup_subject.pack(side="left", padx=(10, 0), pady=2)
entry_followup_subject.bind('<KeyRelease>', lambda e: preview_email() if tree.selection() else None)

followup_message_frame = tk.Frame(followup_frame)
followup_message_frame.pack(fill="x", pady=2)

tk.Label(followup_message_frame, text="Follow-up Message Template:", anchor="nw").pack(side="left", pady=2)
text_followup_message = tk.Text(followup_message_frame, width=45, height=6)
text_followup_message.pack(side="left", padx=(10, 0), pady=2)
text_followup_message.bind('<KeyRelease>', lambda e: root.after_idle(lambda: preview_email() if tree.selection() else None))

# Control buttons and status
frame_controls = tk.Frame(frame_left)
frame_controls.pack(fill="x", pady=10)

# Send/Stop buttons
button_frame = tk.Frame(frame_controls)
button_frame.pack(fill="x", pady=5)

send_button = tk.Button(button_frame, text="Send New Emails", command=send_emails_thread,
                       bg="green", fg="white", font=("Arial", 10, "bold"))
send_button.pack(side="left", padx=5)

stop_button = tk.Button(button_frame, text="Stop Sending", command=stop_sending,
                       bg="red", fg="white", font=("Arial", 10, "bold"))
stop_button.pack(side="left", padx=5)

# Status and countdown
status_frame = tk.Frame(frame_controls)
status_frame.pack(fill="x", pady=5)

status_label = tk.Label(status_frame, text="Ready to send emails",
                       font=("Arial", 9), anchor="w", bg="lightgray", relief="sunken")
status_label.pack(fill="x", pady=2)

countdown_label = tk.Label(status_frame, text="", font=("Arial", 9, "bold"), anchor="w")
countdown_label.pack(fill="x", pady=1)

# Progress bar
progress_frame = tk.Frame(frame_controls)
progress_frame.pack(fill="x", pady=5)

progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(progress_frame, variable=progress_var, maximum=100)
progress_bar.pack(fill="x", pady=2)

progress_label = tk.Label(progress_frame, text="Progress: 0/0 emails sent",
                         font=("Arial", 9), anchor="w")
progress_label.pack(fill="x", pady=1)

# Excel data display
frame_bottom = tk.Frame(frame_left)
frame_bottom.pack(fill="both", expand=True, pady=5)

tk.Label(frame_bottom, text="Excel Data:").pack()
tree = ttk.Treeview(frame_bottom)
tree.pack(fill="both", expand=True)
tree.bind("<<TreeviewSelect>>", show_preview)

# Email preview
preview_header_frame = tk.Frame(frame_right)
preview_header_frame.pack(fill="x", pady=5)

tk.Label(preview_header_frame, text="Email Preview:", font=("Arial", 12, "bold")).pack(side="left")

# Thread status indicator
thread_status_label = tk.Label(preview_header_frame, text="", font=("Arial", 9), fg="blue")
thread_status_label.pack(side="right")

html_preview = HtmlFrame(frame_right, horizontal_scrollbar="auto")
html_preview.pack(fill="both", expand=True)

# Initialize dropdowns and global variables
follow_up_mode = False

# Initialize dropdowns when the app starts
populate_email_dropdown()
populate_time_dropdown()

# Refresh thread status on startup
# refresh_thread_status()  # Removed or comment out since function is not defined

root.mainloop()