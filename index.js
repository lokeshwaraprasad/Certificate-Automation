const XLSX = require("xlsx");
const nodemailer = require("nodemailer");
const fs = require("fs");
const path = require("path");
const { PDFDocument, rgb, StandardFonts } = require("pdf-lib");

// Load Excel File
const filePath = path.join(__dirname, "devscertificates.xlsx");
if (!fs.existsSync(filePath)) {
    console.error(`❌ Excel file not found: ${filePath}`);
    process.exit(1);
}

const workbook = XLSX.readFile(filePath);
const sheet_name = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheet_name];
const data = XLSX.utils.sheet_to_json(sheet);

// Function to Generate Certificate
async function modifyCertificate(name, eventName) {
    try {
        if (!name || !eventName) throw new Error("Invalid Name or Event Name");

        const uppercaseName = name.toUpperCase();
        const uppercaseEvent = eventName.toUpperCase();

        // Load Certificate Template
        const templatePath = "C:\\Users\\vlpvg\\OneDrive\\Desktop\\devs-cer\\templates\\certficate.pdf";
        if (!fs.existsSync(templatePath)) {
            console.error(`❌ Certificate template not found: ${templatePath}`);
            process.exit(1);
        }
        const templateBytes = fs.readFileSync(templatePath);
        const pdfDoc = await PDFDocument.load(templateBytes);
        const pages = pdfDoc.getPages();
        const firstPage = pages[0];
        const pageWidth = firstPage.getWidth();

        // Embed Helvetica-Bold font
        const font = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
        
        // Set font sizes
        const fontSize = 17.5;
        const eventFontSize = 20;

        // Positioning Name
        const nameY = 270;
        const textWidth = font.widthOfTextAtSize(uppercaseName, fontSize);
        const namePadding = 20;
        const nameX = (pageWidth - textWidth) / 2 - namePadding;
        firstPage.drawText(uppercaseName, {
            x: nameX,
            y: nameY,
            size: fontSize,
            font: font,
            color: rgb(0, 0, 0),
        });

        // Positioning Event Name
        const eventPadding = 8;
        const eventY = 230 + eventPadding;
        const eventWidth = font.widthOfTextAtSize(uppercaseEvent, eventFontSize);
        const eventX = (pageWidth - eventWidth) / 2 - 69 + eventPadding;
        firstPage.drawText(uppercaseEvent, {
            x: eventX,
            y: eventY,
            size: eventFontSize,
            font: font,
            color: rgb(0, 0, 0),
        });

        // Save Modified PDF
        const certDir = path.join(__dirname, "certificates");
        if (!fs.existsSync(certDir)) fs.mkdirSync(certDir);
        const outputPath = path.join(certDir, `${name}_certificate.pdf`);
        fs.writeFileSync(outputPath, await pdfDoc.save());

        return outputPath;
    } catch (error) {
        console.error(`❌ Error modifying certificate for ${name}:`, error);
        return null;
    }
}

// Function to Send Email
async function sendEmail(toEmail, name, eventName, certificatePath) {
    let transporter = nodemailer.createTransport({
        service: "gmail",
        auth: {
            user: "devsrec@rajalakshmi.edu.in",
            pass: "xajb basc cibn ieqf",
        },
        tls: { rejectUnauthorized: false },
        connectionTimeout: 10000,
    });

    let mailOptions = {
        from: "devsrec@rajalakshmi.edu.in",
        to: toEmail,
        subject: `Congratulations on Your Participation${eventName}`,
        html: `
            <p style="color: black;">Dear ${name.toUpperCase()},</p>
            <p style="color: black;">Congratulations! Attached is your certificate of participation for <strong>${eventName}</strong> by <strong>DEVS REC</strong>. Your Enthusiasm and efforts truly made the event a Success</p>
            <p style="color: black;">Keep learning,Keep Coding!</p>
            <p style="color: black;">Best regards,</p>
            <p style="color: black;"><strong>DEVS REC Team</strong></p>
        `,
        attachments: [
            {
                filename: `${name}_certificate.pdf`,
                path: certificatePath,
            },
        ],
    };

    try {
        await transporter.sendMail(mailOptions);
        console.log(`✅ Certificate sent to ${name} at ${toEmail}`);
    } catch (error) {
        console.error(`❌ Failed to send email to ${toEmail}:`, error);
        if (error.code === "ESOCKET" || error.code === "ECONNRESET") {
            console.log(`🔄 Retrying to send email to ${toEmail}...`);
            try {
                await transporter.sendMail(mailOptions);
                console.log(`✅ Retry successful: Certificate sent to ${name} at ${toEmail}`);
            } catch (retryError) {
                console.error(`❌ Retry failed for ${toEmail}:`, retryError);
            }
        }
    }
}

// Main Function to Process Certificates
async function processCertificates() {
    for (let person of data) {
        let name = person.Name;
        let email = person.Email;
        let eventName = person["Event Name"];

        let certificatePath = await modifyCertificate(name, eventName);
        if (certificatePath) {
            await sendEmail(email, name, eventName, certificatePath);
        }
    }
}

processCertificates().catch(console.error);
