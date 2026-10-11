pip install pandas

import pandas as pd
import json

# Define the email data structured for a clean layout
emails_data = [
    {
        "id": 1,
        "type": "Phishing (Fake Instagram Alert)",
        "sender": "Security Team <security-alert@lnstagram-support.net>",
        "subject": "Urgent: Unusual login attempt detected",
        "body": "We detected a login attempt from a new device in Moscow, Russia. If this wasn't you, please immediately secure your account by verifying your password here: http://secure-lnstagram.com. Failure to verify within 24 hours will result in permanent account suspension.",
        "clues": ["Look closely at the sender address: 'lnstagram' uses a lowercase 'L' instead of an 'I'.", "The URL uses http:// instead of https:// and links to a non-official domain.", "The message creates false urgency by threatening account suspension."]
    },
    {
        "id": 2,
        "type": "Legitimate (Real Amazon Receipt)",
        "sender": "Amazon.com Orders <auto-confirm@amazon.com>",
        "subject": "Your Amazon.com order confirmation #403-1928374-8291023",
        "body": "Thank you for your shopping! Your order for 'Wireless Bluetooth Headphones' has been confirmed. We will send a notification when your package ships. You can track your order status or view your digital invoice securely inside your standard Amazon Mobile App dashboard at any time.",
        "clues": ["The sender domain is explicitly '@amazon.com' with no misspellings.", "The email references a standard tracking practice instead of forcing you to click an external, suspicious link."]
    },
    {
        "id": 3,
        "type": "Phishing (Fake Teacher Grade Update)",
        "sender": "School Portal <notifications@portal-grades-edu.org>",
        "subject": "CRITICAL: Grade Drop Notice for Quarter 2",
        "body": "Hello Student, your overall grade in Advanced Algebra has dropped below passing threshold. Mr. Davis has updated the portal with critical feedback files. You must download and sign the attached progress report to avoid academic probation: http://schoolportal-login-verify.com",
        "clues": ["The download link ends in '.exe', which is a dangerous executable file, not a document or PDF.", "Schools typically host dashboards on official district domains (.edu or regional K12 domains), not generic '.org' combinations."]
    },
    {
        "id": 4,
        "type": "Phisihing (Fake Delivery Update)",
        "sender": "FedEx Express Delivery <delivery-update-tracking@fedx-post-service.com>",
        "subject": "Action Required: Package address correction needed",
        "body": "Your shipment tracking ID #US-82910-8291 could not be delivered due to an incomplete apartment number. A holding fee of $1.50 is required to reschedule delivery. Click here to update your address details and process payment: https://fedx-package-redirection.net",
        "clues": ["The company name is misspelled in the sender address and link ('fedx' instead of 'fedex').", "Major delivery services do not demand random credit card processing fees via text or unsolicited emails to update addresses."]
    }
]

df = pd.DataFrame(emails_data)
# Displaying structure clearly
print(df[['type', 'sender', 'subject']].to_string(index=False))
