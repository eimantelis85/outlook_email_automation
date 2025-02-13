# **Outlook Email Sending Automation Tool**  

## **Overview**  
Tired of sending the same email individually to multiple recipients? This tool automates the process, allowing you to send personalized emails separately to each recipient using Outlook.  

---

## **How It Works**  
This program reads recipient names and emails from a `list.xlsx` file and sends customized emails to each person individually.  

---

## **Prerequisites**  
Before running the tool, ensure you have:  
✅ **Python** installed on your machine.  

---

## **Getting Started**  

### **1. Clone the Repository**  
```sh
git clone https://github.com/eimantelis85/outlook_email_automation.git
cd outlook_email_automation
```

### **2. Install Dependencies**  
```sh
pip install -r requirements.txt
```

### **3. Configure Environment Variables**  
1. Rename `.env.example` to `.env`:  
   ```sh
   mv .env.example .env
   ```
2. Open `.env` in a text editor and update the following values:  
   - **Outlook credentials**: `OUTLOOK_EMAIL` and `OUTLOOK_PASSWORD`  
   - **CC Email (optional)**: `CC_EMAIL` (delete if not needed)  
   - **Your name (for email signature)**: `YOUR_NAME`  

### **4. Prepare the Recipient List**  
- Open `list.xlsx` and add recipient details:  
  - **Column "Name"** → Enter recipient names.  
  - **Column "Email"** → Enter their email addresses.  

### **5. Customize Your Email Content**  
- Open `main.py` and edit the email text to suit your needs.  

### **6. Start Sending Emails**  
Run the script with:  
```sh
python main.py
```

---

## **Author**  
👤 **Eimantas Venslovas**  
