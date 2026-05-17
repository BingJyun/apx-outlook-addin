# Notes for Testers — APX.AI Outlook Add-in

This document provides all required information for testing the APX.AI Outlook Add-in, including test accounts, server credentials, and step-by-step instructions. Please follow all steps in order. Failure to complete the prerequisites may result in an incomplete test experience.

---

## Prerequisites — Account Setup (Required Before Testing)

Before testing the add-in, both test accounts must be individually activated on the APX.AI platform. This step **cannot be completed on behalf of the reviewer**, as each user must generate and securely store their own private key — this is a core security principle of the product.

**Platform URL:** https://apxpoc.ioneit.com/

### Test Accounts

| Role     | Username          | Password        | Email                          |
|----------|-------------------|-----------------|--------------------------------|
| Sender   | `<sender_username>`   | `<sender_password>`   | `<sender_email@test.com>`   |
| Receiver | `<receiver_username>` | `<receiver_password>` | `<receiver_email@test.com>` |

> Replace the placeholders above with the actual test account credentials provided separately.

### Account Activation Steps (perform for both accounts)

1. Log in at https://apxpoc.ioneit.com/ using the credentials above.
2. Click **Private Key Settings** in the top-right corner.
3. Set a private key password and **download and save the private key file** to your local machine.
4. Repeat steps 1–3 for the second account.

> Both accounts must complete activation before proceeding.

---

## Test Flow 1 — Send a Secure File via Outlook

Use the **Sender** account to perform the following steps:

**Step 1 — Compose a new email in Outlook**
- Set the recipient to the Receiver's email address.

**Step 2 — Open the APX.AI Add-in panel**
- Launch the add-in from the Outlook toolbar or ribbon.

**Step 3 — Configure the server URL**
- Enter: `https://apxpoc.ioneit.com`
- Click **Continue**.

**Step 4 — Log in to the add-in**
- Enter the Sender's username and password.

**Step 5 — Verify your private key**
- Upload the private key file saved during account activation.
- Enter your private key password.

**Step 6 — Select and upload a file**
- The add-in will automatically detect the recipient.
- Select a file to send.
- Click **Upload and Generate Link**.

**Step 7 — Send the email**
- The file name and secure download link will be automatically inserted into the email body.
- Complete and send the email.

---

## Test Flow 2 — Receive and Download the File

Use the **Receiver** account to perform the following steps:

1. Open the received email and click the secure download link.
2. You will be directed to: https://apxpoc.ioneit.com/
3. Log in with the Receiver's credentials.
4. Click **Received Files** in the left sidebar.
5. Select the file and click download.

The file will be decrypted using the receiver's private key and downloaded successfully.

**This completes the full test flow.**

---

> For any issues during testing, please contact: **support@ioneit.com**
