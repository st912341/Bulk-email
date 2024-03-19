const nodemailer = require('nodemailer');
const xlsx = require('xlsx');
const fs = require('fs');

// Read the Excel file
const workbook = xlsx.readFile('leads.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = xlsx.utils.sheet_to_json(sheet);

// Email configuration
const transporter = nodemailer.createTransport({
    host: 'smtp.hostinger.com', // Adjust this to your SMTP server
    port: 587, // SMTP port (usually 587 for TLS)
    secure: false, // true for 465, false for other ports
    auth: {
        user: 'info@webgisolutions.com',
        pass: 'Atewhspeed#95076'
    }
});

// Initialize counters
let totalSent = 0;
let totalDelivered = 0;
let totalFailed = 0;

// Create a follow-up sheet
let followUpData = [];

// Send emails
data.forEach((lead, index) => {
    const mailOptions = {
        from: '"WebGi Solutions" <info@webgisolutions.com>',
        to: lead.email,
        subject: 'Common Subject for All Emails',
        html: `<p>Dear ${lead.name},</p><p>${lead.website}</p>`
    };

    transporter.sendMail(mailOptions, (error, info) => {
        if (error) {
            console.log(`Error sending email to ${lead.email}: ${error}`);
            totalFailed++;
            followUpData.push({ Email: lead.email, Status: 'Failed' });
        } else {
            console.log(`Email sent to ${lead.email}: ${info.messageId}`);
            totalDelivered++;
            followUpData.push({ Email: lead.email, Status: 'Delivered' });
        }

        // Update counters
        totalSent++;

        // If all emails have been processed, update the Excel file
        if (totalSent === data.length) {
            console.log(`Total Sent: ${totalSent}`);
            console.log(`Total Delivered: ${totalDelivered}`);
            console.log(`Total Failed: ${totalFailed}`);

            // Add report summary to follow-up data
            followUpData.push({ Email: 'Total Sent', Status: totalSent });
            followUpData.push({ Email: 'Total Delivered', Status: totalDelivered });
            followUpData.push({ Email: 'Total Failed', Status: totalFailed });

            // Add follow-up sheet to workbook
            const followUpSheet = xlsx.utils.json_to_sheet(followUpData);
            xlsx.utils.book_append_sheet(workbook, followUpSheet, 'Follow Up');

            // Write updated workbook to file
            xlsx.writeFile(workbook, 'leads.xlsx');
        }
    });
});
