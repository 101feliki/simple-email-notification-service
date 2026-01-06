const XLSX = require("xlsx");
const nodemailer = require("nodemailer");
const fs = require("fs");
const path = require("path");

console.log("🚀 BIRDVIEW BULK EMAIL SENDER");
console.log("=============================\n");


const config = {
      SMTP_HOST: process.env.SMTP_HOST,
    SMTP_PORT: parseInt(process.env.SMTP_PORT) || 465,
    SMTP_USER: process.env.SMTP_USER,
    SMTP_PASS: process.env.SMTP_PASS, // Now loaded from .env
    FROM_EMAIL: process.env.FROM_EMAIL,
    FROM_NAME: process.env.FROM_NAME,
    SUBJECT: process.env.EMAIL_SUBJECT || "Inquiry of Unfinished Insurance Purchase",
    DELAY_MS: parseInt(process.env.EMAIL_DELAY_MS) || 15000,
    BATCH_SIZE: parseInt(process.env.BATCH_SIZE) || 10,
    BATCH_DELAY: parseInt(process.env.BATCH_DELAY_MS) || 3600000,
};

// Function to format date like "Thursday, December 18, 2025 11:32 AM"
function formatEmailDate() {
    const now = new Date();
    const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    
    const dayName = days[now.getDay()];
    const monthName = months[now.getMonth()];
    const day = now.getDate();
    const year = now.getFullYear();
    
    let hours = now.getHours();
    const minutes = now.getMinutes().toString().padStart(2, '0');
    const ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    hours = hours ? hours : 12;
    
    return `${dayName}, ${monthName} ${day}, ${year} ${hours}:${minutes} ${ampm}`;
}

// Function to clean product name
function cleanProductName(product) {
    if (!product) return "Insurance";
    
    let cleaned = product
        .replace(/Personl Accident/gi, "Personal Accident")
        .replace(/Evacuation\s*&\s*Repatriation/gi, "Evacuation and Repatriation")
        .trim();
    
    // Add "Cover" if medical and doesn't have it
    if (cleaned.toLowerCase().includes('medical') && !cleaned.toLowerCase().includes('cover')) {
        cleaned = cleaned + ' Cover';
    }
    
    return cleaned;
}

// Function to generate HTML email template
function generateEmailTemplate(client, emailDate) {
    return `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        body {
            margin: 0;
            padding: 0;
            font-family: Arial, sans-serif;
            background-color: #ffffff;
            line-height: 1.6;
            color: #000000;
        }
    </style>
</head>
<body>
    <div style="width: 100%; padding: 20px; box-sizing: border-box; max-width: 800px; margin: 0 auto;">
        
        <!-- Email Header (simulating email client display) -->
        <div style="font-size: 12px; color: #666; padding-bottom: 15px; border-bottom: 1px solid #e0e0e0; margin-bottom: 20px;">
            <div><strong>Birdview Customer Care &lt;customerservice@birdviewinsurance.com&gt;</strong></div>
            <div>${emailDate}</div>
            <div>to me</div>
        </div>

        <!-- Main Email Content -->
        <div style="font-size: 14px; color: #000000;">
            <p>Dear ${client.name},</p>
            
            <p>&nbsp;</p>
            
            <p>Thank you for visiting the Birdview Insurance portal and showing interest in our insurance solutions.</p>
            
            <p>We noticed that you started the purchase process for a ${client.product} but were unable to complete it. We understand that sometimes challenges may arise, and we would really appreciate the opportunity to learn from your experience.</p>
            
            <p>Kindly let us know:</p>
            
            <p>Did you encounter any difficulty on the portal?<br>
            Was there any information that was unclear?<br>
            Is there any way our team can assist you to complete your purchase?</p>
            
            <p>Our team is ready to support you and guide you through the process to ensure you get the cover that best suits your needs.</p>
            
            <p>You may reply to this email or reach us directly on <strong>+254 111 056 610</strong> for immediate assistance.</p>
            
            <p>Thank you for considering Birdview Insurance. We look forward to serving you.</p>
            
            <p>&nbsp;</p>
            <p>&nbsp;</p>
            <p>&nbsp;</p>
            
            <!-- Customer Service Team -->
            <div style="margin-top: 20px;">
                <p style="margin: 0;"><strong>Customer Service Team</strong><br>
                +254 111 056 610</p>
            </div>
            <div style="text-align: left; margin-top: 13px; padding-top: 15px; border-top: 1px solid #e0e0e0;">
                <a href="https://www.birdviewmicroinsurance.com" target="_blank" style="text-decoration: none;">
                    <img width="120" src="https://www.birdviewmicroinsurance.com/_next/image?url=%2Fimages%2Flogo.jpeg&w=384&q=75" alt="Birdview Microinsurance" style="border-radius: 5px; vertical-align: middle;">
                </a>
                <div style="margin-top: 10px; font-size: 12px; color: #666;">
                    <a href="https://x.com/BirdviewInsur" target="_blank" style="color: #666; text-decoration: none; margin: 0 5px;">Twitter</a> | 
                    <a href="https://www.facebook.com/BirdviewMicroinsurance" target="_blank" style="color: #666; text-decoration: none; margin: 0 5px;">Facebook</a> | 
                    <a href="https://www.linkedin.com/in/birdview-microinsurance-71b32431a" target="_blank" style="color: #666; text-decoration: none; margin: 0 5px;">LinkedIn</a>
                </div>
            </div>
            
            <p>&nbsp;</p>
            <p>&nbsp;</p>
            
            <!-- Contact Information -->
            <div style="margin-top: 8px;">
                <p style="margin: 8px 0;color: #157EBC"><strong>Email Us:</strong> customerservice@birdviewinsurance.com</p>
                <p style="margin: 8px 0;color: #0f0f10ff"><strong>Call Us:</strong> +254 742 222 888</p>
            </div>
            
            <p>&nbsp;</p>
            
            <!-- Products -->
            <div style="text-align:left; font-size: 14px; color: #157EBC; font-weight: bold; margin: 8px 0; padding: 10px 0; border-top: 1px solid #e0e0e0; border-bottom: 1px solid #e0e0e0;">
                <span>Evacuation and Repatriation | Last Expense | Medical | Hospital Cash | Personal Accident</span>
            </div>
            
            <!-- Disclaimer -->
            <div style="background-color: #f5f5f5; padding: 15px; font-size: 11px; color: #666; text-align: justify; margin-top: 20px; border-radius: 3px;">
                <p style="margin: 0; line-height: 1.5;">
                    DISCLAIMER: This email, along with any attachments, is confidential and intended solely for the use of the individual or entity to whom it is addressed. If you have received this email in error, please notify our system administrator immediately. The content of this email is proprietary to Birdview Microinsurance Limited and strictly confidential. If you are not the intended recipient, any disclosure, copying, distribution, or other use of the information contained in this email is strictly prohibited. Please delete this email from your system and notify the sender if you have received it by mistake.
                </p>
            </div>
            
            <!-- Optional: Add logo at the bottom -->
            
        </div>
    </div>
</body>
</html>`;
}

// Function to generate plain text email
function generateTextEmail(client, emailDate) {
    return `Birdview Customer Care <customerservice@birdviewinsurance.com>
${emailDate}
to me

Dear ${client.name},

Thank you for visiting the Birdview Insurance portal and showing interest in our insurance solutions.

We noticed that you started the purchase process for a ${client.product} but were unable to complete it. We understand that sometimes challenges may arise, and we would really appreciate the opportunity to learn from your experience.

Kindly let us know:

Did you encounter any difficulty on the portal?
Was there any information that was unclear?
Is there any way our team can assist you to complete your purchase?

Our team is ready to support you and guide you through the process to ensure you get the cover that best suits your needs.

You may reply to this email or reach us directly on +254 111 056 610 for immediate assistance.

Thank you for considering Birdview Insurance. We look forward to serving you.



Customer Service Team
+254 111 056 610



Email Us: customerservice@birdviewinsurance.com

Call Us: +254 742 222 888


Evacuation and Repatriation | Last Expense | Medical | Hospital Cash | Personal Accident

DISCLAIMER: This email, along with any attachments, is confidential and intended solely for the use of the individual or entity to whom it is addressed. If you have received this email in error, please notify our system administrator immediately. The content of this email is proprietary to Birdview Microinsurance Limited and strictly confidential. If you are not the intended recipient, any disclosure, copying, distribution, or other use of the information contained in this email is strictly prohibited. Please delete this email from your system and notify the sender if you have received it by mistake.`;
}

async function main() {
    console.log("1. Looking for Excel files...");
    
    // Find Excel files
    const files = fs.readdirSync("./data").filter(file => 
        file.toLowerCase().endsWith(".xlsx") || 
        file.toLowerCase().endsWith(".xls")
    );
    
    if (files.length === 0) {
        console.error("❌ No Excel files found in ./data folder");
        console.log("   Please place your Excel file (test 1.xlsx) in the 'data' folder");
        return;
    }
    
    const excelFile = `./data/${files[0]}`;
    console.log(`✅ Found: ${files[0]}`);
    
    console.log("\n2. Reading Excel data...");
    
    // Read Excel
    const workbook = XLSX.readFile(excelFile);
    const sheetName = workbook.SheetNames[0];
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
    
    // Extract clients
    const clients = [];
    let skipped = 0;
    
    data.forEach((row, index) => {
        // Try different column names
        const name = row["Name"] || row["name"] || row["Client Name"] || "";
        const email = row["Email"] || row["email"] || row["Email Address"] || "";
        let product = row["Product"] || row["product"] || row["Products"] || "";
        
        // Clean product names
        if (product) {
            product = cleanProductName(product.toString());
        }
        
        if (email && email.includes("@") && name) {
            clients.push({
                name: name.toString().trim(),
                email: email.toString().trim().toLowerCase(),
                product: product || "Insurance Cover",
                row: index + 2
            });
        } else {
            skipped++;
        }
    });
    
    console.log(`✅ Found ${clients.length} valid clients (skipped ${skipped} rows)`);
    
    if (clients.length === 0) {
        console.error("❌ No valid clients found");
        return;
    }
    
    // Show sample
    console.log("\n📋 SAMPLE CLIENTS:");
    clients.slice(0, 5).forEach((client, i) => {
        console.log(`${i + 1}. ${client.name} - ${client.email}`);
        console.log(`   Product: ${client.product}`);
    });
    if (clients.length > 5) {
        console.log(`... and ${clients.length - 5} more`);
    }
    
    console.log("\n3. Testing email connection...");
    
    const transporter = nodemailer.createTransport({
        host: config.SMTP_HOST,
        port: config.SMTP_PORT,
        secure: true,
        auth: {
            user: config.SMTP_USER,
            pass: config.SMTP_PASS
        }
    });
    
    try {
        await transporter.verify();
        console.log("✅ Email server connected successfully");
    } catch (error) {
        console.error("❌ Email connection failed:", error.message);
        console.log("\n⚠️  Please check:");
        console.log("1. Is the password correct?");
        console.log("2. Is SMTP server accessible?");
        return;
    }
    
    console.log("\n4. Ready to send emails!");
    console.log(`   Will send to ${clients.length} clients`);
    console.log(`   Delay between emails: ${config.DELAY_MS}ms`);
    
    // Ask for confirmation
    const readline = require("readline");
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });
    
    const answer = await new Promise(resolve => {
        rl.question("\nProceed with sending? (yes/no): ", resolve);
    });
    rl.close();
    
    if (answer.toLowerCase() !== "yes" && answer.toLowerCase() !== "y") {
        console.log("❌ Operation cancelled");
        return;
    }
    
    console.log("\n5. Sending emails...\n");
    
    let success = 0;
    let failed = 0;
    
    for (let i = 0; i < clients.length; i++) {
        const client = clients[i];
        
        console.log(`[${i + 1}/${clients.length}] ${client.name} <${client.email}>`);
        
        const emailDate = formatEmailDate();
        
        try {
            // Prepare mail options
            const mailOptions = {
                from: `"${config.FROM_NAME}" <${config.FROM_EMAIL}>`,
                to: client.email,
                subject: config.SUBJECT,
                html: generateEmailTemplate(client, emailDate),
                text: generateTextEmail(client, emailDate)
            };
            
            await transporter.sendMail(mailOptions);
            
            console.log("   ✅ Email sent");
            success++;
            
        } catch (error) {
            console.error(`   ❌ Failed: ${error.message}`);
            failed++;
        }
        
        // Wait between emails
        if (i < clients.length - 1) {
            console.log(`   ⏳ Waiting ${config.DELAY_MS}ms...\n`);
            await new Promise(resolve => setTimeout(resolve, config.DELAY_MS));
        }
    }
    
    // Show results
    console.log("\n" + "=".repeat(50));
    console.log("📊 RESULTS");
    console.log("=".repeat(50));
    console.log(`✅ Successful: ${success}`);
    console.log(`❌ Failed: ${failed}`);
    console.log(`📋 Total: ${clients.length}`);
    
    // Save report
    const timestamp = new Date().toISOString().split("T")[0];
    const report = `
BULK EMAIL REPORT
Date: ${new Date().toLocaleString()}
Excel File: ${files[0]}
Total Clients: ${clients.length}
Successful: ${success}
Failed: ${failed}

CLIENT LIST:
${clients.map(c => `${c.name} | ${c.email} | ${c.product}`).join("\n")}
`;
    
    fs.writeFileSync(`email-report-${timestamp}.txt`, report);
    console.log(`\n📄 Report saved to: email-report-${timestamp}.txt`);
}

// Run the program
main().catch(error => {
    console.error("❌ Fatal error:", error.message);
    process.exit(1);
});