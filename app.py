"""
Meeting Follow-up Email Generator
=================================

Internal tool for wealth management team to convert client meeting recordings
into professional follow-up emails.

Setup Instructions:
1. Install dependencies: pip install -r requirements.txt
2. Create a .env file with your credentials (see .env)
3. Run the app: streamlit run app.py
4. Upload audio files and generate professional follow-up emails

Author: Investment Fund AI Team
"""

import streamlit as st
import os
import datetime
import json
from pathlib import Path
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import openai
from dotenv import load_dotenv
import uuid
import time

# Load environment variables
load_dotenv()
print("Loaded OpenAI key:", os.getenv("OPENAI_API_KEY"))



# Configure page
st.set_page_config(
    page_title="Meeting Follow-up Generator",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize directories
AUDIO_DIR = Path("audio_files")
EMAILS_DIR = Path("emails")
AUDIO_DIR.mkdir(exist_ok=True)
EMAILS_DIR.mkdir(exist_ok=True)

# Configure OpenAI
openai.api_key = os.getenv("OPENAI_API_KEY")

class MeetingEmailGenerator:
    def __init__(self):
        self.client = openai.OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

    def transcribe_audio(self, audio_file_path):
        """Transcribe audio using OpenAI Whisper API"""
        try:
            with open(audio_file_path, "rb") as audio_file:
                transcript = self.client.audio.transcriptions.create(
                    model="whisper-1",
                    file=audio_file,
                    response_format="text"
                )
            return transcript
        except Exception as e:
            st.error(f"Transcription failed: {str(e)}")
            return None

    def generate_email(self, transcript, recipient_name=""):
        """Generate professional follow-up email using GPT-3.5-turbo"""

        prompt = f"""
        You are James Telfser, a professional wealth management advisor at Ewing Morris. Based on the meeting transcript below, generate a professional follow-up email that matches your established tone and format.

        Your email style characteristics:
        - Warm, professional greeting
        - Brief friendly opener acknowledging the meeting
        - Clear structure with sections like "Key Takeaways" and "Next Steps"
        - Use bullet points and numbered lists for clarity
        - Bold formatting for important sections (**text**)
        - Professional but personable tone
        - Include specific details and action items
        - End with warm sign-off like "All the best, James" or "Warm regards, James"

        Example structure to follow:
        - Greeting with recipient name
        - Brief intro acknowledging the meeting
        - **Participants** section
        - **Objective** section  
        - **Key Takeaways** with numbered points and sub-bullets
        - **Next Steps** or **Action Items**
        - Professional closing

        Meeting Transcript:
        {transcript}

        Generate a professional follow-up email that summarizes the key discussion points, action items, and next steps. Make it client-friendly and well-formatted.Below are a few examples of emails I have already sent. Try and replicate this style to your best ability:
        
        Hi Aidan,

 

Thanks for the time today – exciting times for you!  Keep it going.

 

I thought it would be helpful to follow up with some notes to help with your understanding and to make sure we are on the same page.  

 

Please review and provide feedback, as necessary. 

 

-------------

Participants: Aidan Harradence and James Telfser via Teams Meeting.

Objective: 1) Get caught up ; 2) Review Portfolio; 3) Discuss changes to the portfolio based on market conditions & current lifestyle.

 

Key Takeaways:

Investment Accounts:
Performance has improved this year with accounts up ~20% after fees.  
It’s been a great year for small/mid cap stocks, and we were waiting on this rebound to rebalance the account which is great (ie. we did not sell at the lows).
We have been raising some cash ($20k) throughout the year to diversify the account into less volatile investments.
We talked about ETFs and how we will be utilizing these more going forward to take advantage of “time arbitrage”.
Planning & Personal:
You are currently living in the UK and are pursuing an entrepreneurial venture in the film industry (very cool!).  
You currently have ~5-6 months of living expenses covered assuming you do take on any part time work.
You do not expect to require any of these savings, but would like to keep some in cash as a back up (~$20k).
You reiterated that you are still a beginner when it comes to investing and would like some book recommendations (we have these all in the office next time you are here:
A Random Walk Down Wall Street
Rich Dad Poor Dad
The Psychology of Money
Reporting:
You would like a login for NDEX/NBIN+
You would like to set up a recurring quarterly check in, especially given that you are trying to learn and may require additional income.
Next Steps:

Action Items:
You will review what your CRA details say online about what you can contribute to your TFSA.
James to send an email with an updated asset allocation and changes to the accounts incorporating your updated information and desire for less volatility.
Katherine will send over login details for our systems so you can track the accounts.
 

All the best,
James

Heres another-

Hi everyone,


It was great to see you all yesterday.  Please see below for minutes and some additional details. 

 

As always, please let me know if I have missed anything or if you have any questions.


All the best,
James

 

Participants: Brad Jarvis (BJ), Faye Wardrop (FW), Steve Barlow (SB) and James Telfser (JT) – Victoria in person

Objective: 1) Quarterly Account Review;  2) Market Outlook

 

Key Takeaways:

Quarterly Review:
We discussed performance, asset allocation and recent changes.  Overall, Jarvis Group of accounts was +1% in Q1-25.
Major contributors were gold, global equities ex-US, Canadian large cap equities and preferred shares.
Weakness was driven by the large cap US Equity ETFs as technology/consumer discretionary underperformed.
We remained within the Target Asset Allocation ranges as per our recent IPS.
We discussed the changes to the Aventine Funds – Selling down the US Equity Fund (for more passive and fixed income exposure) & an eventual merger of the Aventine Canadian Equity Fund.
We currently have orders in to sell $200k ACE Fund and $450k of the US Equity Fund at April 30th, representing 4.5% additional cash. 
We discussed the NorthStream Credit Fund sell which settles early May ($425,000 or ~3% of the account)
We will hold a higher than normal cash position in the account, however we will add to fixed income & infrastructure (KKR) to round out the asset allocation. 
We discussed continuing to hold a higher percentage of cash/fixed income/gold which we all agreed was a good idea as we deal with the uncertainty at the moment.
We agreed that holding cash/money market, despite lower rates in CAD was the right decision given the volatility in currencies and goal of capital preservation, especially for downpayment cash.
IPS/KYC:
BJ has an offer on Saturna ($1.34mm less expenses) and is working towards a close in the coming months.
BJ also has an offer on one of the lots for $300k less expenses.
It is currently expected that BJ will require $1-$1.2mm in cash for downpayment on Salt Spring in addition to cash from the properties above (sometime in next 12 months)
JT to continue to be conservative with cash and keep enough liquidity for downpayment (currently $2.8mm in cash). 
Upon the withdrawal of the funds for downpayment, JT to rebalance to ensure that we remain in line with IPS targets and current market outlook.
Expenses for BJ seem to be a little higher than originally projected in our IPS from a couple years ago.
Next Steps:
JT to make asset allocation changes in Q2 once aforementioned trades settle in early May. 
JT to continue to build up the passive side of the accounts with proceeds from active mandates, however, will be more conservative than normal given current expectation for volatility and potential cash needs.
JT to include FW on any asset allocation or special investment opportunities communications. 
FW will ensure that the appropriate accounting/legal team is in the loop.
JT to follow up with BJ/FW regarding an updated IPS to incorporate new and more realistic spending targets. 
JT to offer up financial planning exercise to ensure current spending is sustainable under conservative assumptions.
BJ/FW to keep JT informed of any structural changes to the accounts or changes to investment criteria or eligibility.
Have a great day ahead!

James

this was more how he would write the email if it was to a 3rd party or a larg client not an individual investor:

heres a final example :

Hi Yvonne,


It was great to see you in person yesterday!  Let’s make sure to do that at least once per year 😊.  Look forward to hearing about your 4 generation(!) trip to NFLD.

 

Please see below for minutes and some additional details. 

 

As always, please let me know if I have missed anything or if you have any questions.


All the best,
James

 

Participants: Yvonne Taylor (YT), James Telfser (JT) – In Person

Objective: 1) YTD Account Review;  2) KYC Check up

 

Key Takeaways:

YTD Review:
We discussed performance, asset allocation and recent changes to the accounts and fund structures. Overall, your accounts are up ~1% in 2025, up just shy of 6% per year after fees since inception.  
Major contributors were preferred shares, Canadian large cap equities and global equities (IQLT ETF).  
Weakness was primarily driven by Apple (-20% YTD + USD weakness) and Transforce (-36% YTD).
We remained within the Target Asset Allocation ranges of 0-30% Cash,  15-35% Fixed Income, and 30-70% Equity.  We’ve kept it simple without allocations to our alternatives sleeve.
We discussed the changes to the Aventine Funds – Selling down the US Equity Fund (for more passive and fixed income exposure) & an eventual merger of the Aventine Canadian Equity Fund.
We currently have orders in to sell all but ~$1,000 of the ACE Fund at the end of June.  
We discussed the EM Flexible Fixed Income Fund and the other EM products.  JT made it clear we will not add additional Ewing Morris products.
We also walked through the conflicts of interest document.
We walked through the various ETFs and the benefits of holdings them (QQQ/HXQ = Tech w/o concentration risk, XIC/TSX = Broad Canada, resource exposure w/o having to pick, etc). 
Finally, we discussed the various fixed income ETFs and why we hold them.  Reasons include higher yield and view that the USD will appreciate vs CAD so could make 4% income and an additional 2-3% in currency appreciation.
IPS/KYC:
No change in financial position or asset allocation, currently receiving enough cash from the monthly RRIF payments.
Continue to manage account considering time horizon and avoid any illiquid investments.
No major purchases/withdrawals expected. 
Other/Next Steps:
JT to enter sell order for ACE/Darkhorse
JT reviewed TFSA contribution room, and you CAN contribute the $85,000 (deregistered in 2024) in 2025.
JT to contribute the AirBoss and Preferred Share position to shelter from taxes going forward.
JT to send Addepar login and upload presentation
Completed…please let me know if you have trouble logging in.  
Have a great day ahead!

James

        {"If you can identify the client's name from the transcript, use it in the greeting. Otherwise, use a general greeting." if not recipient_name else f"Address the email to {recipient_name}."}
        """

        try:
            response = self.client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system",
                     "content": "You are a professional wealth management advisor who writes clear, structured follow-up emails after client meetings."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=1500,
                temperature=0.7
            )
            return response.choices[0].message.content
        except Exception as e:
            st.error(f"Email generation failed: {str(e)}")
            return None

    def send_email(self, recipient_email, subject, email_body, transcript=None):
        """Send email via SMTP"""
        try:
            # Email configuration
            smtp_server = os.getenv("SMTP_SERVER", "smtp.gmail.com")
            smtp_port = int(os.getenv("SMTP_PORT", "587"))
            sender_email = os.getenv("SENDER_EMAIL")
            sender_password = os.getenv("SENDER_PASSWORD")

            if not all([sender_email, sender_password]):
                st.error("Email credentials not configured. Please check your .env file.")
                return False

            # Create message
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = recipient_email
            msg['Subject'] = subject

            # Add email body
            msg.attach(MIMEText(email_body, 'plain'))

            # Optionally attach transcript
            if transcript and st.session_state.get('include_transcript', False):
                transcript_attachment = MIMEText(transcript, 'plain')
                transcript_attachment.add_header('Content-Disposition', 'attachment', filename='meeting_transcript.txt')
                msg.attach(transcript_attachment)

            # Send email
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(sender_email, sender_password)
            text = msg.as_string()
            server.sendmail(sender_email, recipient_email, text)
            server.quit()

            return True

        except Exception as e:
            st.error(f"Failed to send email: {str(e)}")
            return False

    def save_email_record(self, recipient_email, subject, email_body, transcript, audio_filename):
        """Save email record for future reference"""
        record = {
            "timestamp": datetime.datetime.now().isoformat(),
            "recipient_email": recipient_email,
            "subject": subject,
            "email_body": email_body,
            "transcript": transcript,
            "audio_filename": audio_filename,
            "id": str(uuid.uuid4())
        }

        filename = f"email_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        filepath = EMAILS_DIR / filename

        with open(filepath, 'w') as f:
            json.dump(record, f, indent=2)

        return filepath


def main():
    st.title("📧 Ewing Morris Follow-up Email Generator")
    st.markdown("*Transform client meeting recordings into professional follow-up emails*")

    # Initialize session state
    if 'generator' not in st.session_state:
        st.session_state.generator = MeetingEmailGenerator()

    # Sidebar for settings and history
    with st.sidebar:
        st.header("⚙️ Settings")

        # API Key check
        if not os.getenv("OPENAI_API_KEY"):
            st.error("⚠️ OpenAI API key not found. Please configure your .env file.")
        else:
            st.success("✅ OpenAI API configured")

        if not all([os.getenv("SENDER_EMAIL"), os.getenv("SENDER_PASSWORD")]):
            st.warning("⚠️ Email credentials not configured. Email sending will be disabled.")
        else:
            st.success("✅ Email credentials configured")

        st.markdown("---")

        # Email options
        st.subheader("Email Options")
        include_transcript = st.checkbox("Include transcript as attachment", value=False)
        st.session_state.include_transcript = include_transcript

        st.markdown("---")

        # Recent emails
        st.subheader("📁 Recent Emails")
        email_files = list(EMAILS_DIR.glob("*.json"))
        if email_files:
            email_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
            for email_file in email_files[:5]:  # Show last 5
                with open(email_file, 'r') as f:
                    record = json.load(f)
                st.text(f"{record['timestamp'][:16]}")
                st.text(f"To: {record['recipient_email']}")
                st.markdown("---")
        else:
            st.text("No emails generated yet")

    # Main content area
    col1, col2 = st.columns([1, 1])

    with col1:
        st.header("🎤 Upload Audio")

        # File uploader
        uploaded_file = st.file_uploader(
            "Choose an audio file",
            type=['mp3', 'wav', 'm4a', 'flac'],
            help="Upload your client meeting recording"
        )

        # Recipient email input
        recipient_email = st.text_input(
            "Recipient Email",
            placeholder="client@example.com",
            help="Email address to send the follow-up to"
        )

        # Optional recipient name
        recipient_name = st.text_input(
            "Client Name (optional)",
            placeholder="e.g., John Smith",
            help="Client name for personalized greeting"
        )

        # Generate button
        generate_button = st.button(
            "🔄 Generate Email",
            type="primary",
            disabled=not uploaded_file or not recipient_email
        )

    with col2:
        st.header("📄 Results")

        if generate_button and uploaded_file and recipient_email:
            with st.spinner("Processing audio and generating email..."):
                # Save uploaded file
                audio_filename = f"{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}_{uploaded_file.name}"
                audio_path = AUDIO_DIR / audio_filename

                with open(audio_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                # Transcribe audio
                st.info("🎧 Transcribing audio...")
                transcript = st.session_state.generator.transcribe_audio(audio_path)

                if transcript:
                    st.success("✅ Transcription completed!")

                    # Generate email
                    st.info("📝 Generating professional email...")
                    email_body = st.session_state.generator.generate_email(transcript, recipient_name)

                    if email_body:
                        st.success("✅ Email generated!")

                        # Store results in session state
                        st.session_state.current_transcript = transcript
                        st.session_state.current_email = email_body
                        st.session_state.current_recipient = recipient_email
                        st.session_state.current_audio_filename = audio_filename
                        st.session_state.current_audio_path = audio_path

    # Display results if available
    if hasattr(st.session_state, 'current_transcript'):
        st.markdown("---")

        # Create tabs for transcript and email
        tab1, tab2 = st.tabs(["📝 Generated Email", "📄 Full Transcript"])

        with tab1:
            st.subheader("Professional Follow-up Email")

            # Display email in a styled container
            st.markdown(
                f"""
                <div style="
                    background-color: #f8f9fa;
                    padding: 20px;
                    border-radius: 10px;
                    border-left: 4px solid #007acc;
                    font-family: 'Arial', sans-serif;
                    line-height: 1.6;
                ">
                    <pre style="white-space: pre-wrap; font-family: inherit;">{st.session_state.current_email}</pre>
                </div>
                """,
                unsafe_allow_html=True
            )

            # Send email section
            st.markdown("---")
            col1, col2, col3 = st.columns([1, 1, 1])

            with col1:
                if st.button("📧 Send Email", type="primary"):
                    subject = "Follow-Up from Our Recent Meeting"

                    with st.spinner("Sending email..."):
                        success = st.session_state.generator.send_email(
                            st.session_state.current_recipient,
                            subject,
                            st.session_state.current_email,
                            st.session_state.current_transcript
                        )

                    if success:
                        st.success(f"✅ Email sent successfully to {st.session_state.current_recipient}!")

                        # Save record
                        record_path = st.session_state.generator.save_email_record(
                            st.session_state.current_recipient,
                            subject,
                            st.session_state.current_email,
                            st.session_state.current_transcript,
                            st.session_state.current_audio_filename
                        )

                        st.info(f"📁 Email record saved: {record_path.name}")

                        # Clean up audio file
                        if st.session_state.current_audio_path.exists():
                            os.remove(st.session_state.current_audio_path)
                            st.info("🗑️ Audio file cleaned up")

            with col2:
                # Download email as text file
                email_text = f"Subject: Follow-Up from Our Recent Meeting\n\n{st.session_state.current_email}"
                st.download_button(
                    "💾 Download Email",
                    email_text,
                    file_name=f"email_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                    mime="text/plain"
                )

            with col3:
                # Copy to clipboard button (JavaScript required)
                if st.button("📋 Copy Email"):
                    st.info("Email content ready to copy (select and copy from the box above)")

        with tab2:
            st.subheader("Meeting Transcript")
            st.text_area(
                "Full transcript of the meeting:",
                value=st.session_state.current_transcript,
                height=400,
                disabled=True
            )

            # Download transcript
            st.download_button(
                "💾 Download Transcript",
                st.session_state.current_transcript,
                file_name=f"transcript_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                mime="text/plain"
            )

    # Footer
    st.markdown("---")
    st.markdown(
        """
        <div style="text-align: center; color: #666; padding: 20px;">
            <p>Meeting Follow-up Email Generator | Built for Investment Fund AI Team</p>
            <p>Secure • Professional • Efficient</p>
        </div>
        """,
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()# Updated UI - Wed 25 Jun 2025 15:37:05 EDT

