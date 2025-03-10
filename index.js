require('dotenv').config();
const TelegramApi = require('node-telegram-bot-api');
const xlsx = require("xlsx");

const { S3Client, GetObjectCommand, PutObjectCommand } = require("@aws-sdk/client-s3");
const { PassThrough } = require('stream');

const token = process.env.TELEGRAM_BOT_TOKEN;
const bot = new TelegramApi(token, { polling: true });

const s3 = new S3Client({
    endpoint: process.env.S3_ENDPOINT,
    region: "us-east-1",
    credentials: {
        accessKeyId: process.env.S3_ACCESS_KEY_ID,
        secretAccessKey: process.env.S3_SECRET_ACCESS_KEY,
    },
});

async function streamToBuffer(stream) {
    return new Promise((resolve, reject) => {
        const chunks = [];
        stream.on("data", (chunk) => chunks.push(chunk));
        stream.on("end", () => resolve(Buffer.concat(chunks)));
        stream.on("error", reject);
    });
}

async function downloadExcelBuffer(key) {
    const command = new GetObjectCommand({ Bucket: "tables", Key: key });
    const response = await s3.send(command);
    return await streamToBuffer(response.Body);
}

async function uploadBuffer(key, buffer) {
    const params = {
        Bucket: "tables",
        Key: key,
        Body: buffer,
        ContentLength: buffer.length,
    };
    return await s3.send(new PutObjectCommand(params));
}

const allowedUsers = [7540947010, 7529522452, 7649862662];
let addingMode = false;
let command = '';

bot.setMyCommands([
    { command: "/start", description: "Start the bot and see options" },
    { command: "/in", description: "Add 'in' row to a table" },
    { command: "/out", description: "Add 'out' row to a table" },
    { command: "/undo", description: "Delete the latest entry in the table" },
    { command: "/list", description: "Retrieve the latest version of the table" }
]);

const isUserAllowed = (id) => allowedUsers.includes(id);

bot.onText(/\/start/, (msg) => {
    if (!isUserAllowed(msg.from.id)) {
        bot.sendMessage(msg.chat.id, "🚫 Access denied. You are not authorized to use this bot.");
        return;
    }
    bot.sendMessage(msg.chat.id, `Hello ${msg.from.first_name}! 😊
This bot updates users' tables stored in Storj.

📌 Commands:
• /start - Info.
• /in - Add an 'in' row to the user's table.
• /out - Add an 'out' row to the user's table.
• /undo - Delete the latest entry from the table.
• /list - Retrieve the latest version of the table.

After sending /in or /out, please send the parameters as:
• For /in: {description} {amount} {percent}
• For /out: {description} {amount}`);
});

bot.onText(/^\/in(@[\w_]+)?$/, (msg) => {
    if (!isUserAllowed(msg.from.id)) {
        bot.sendMessage(msg.chat.id, "🚫 Access denied. You are not authorized to use this bot.");
        return;
    }
    addingMode = true;
    command = 'in';
});

bot.onText(/^\/out(@[\w_]+)?$/, (msg) => {
    if (!isUserAllowed(msg.from.id)) {
        bot.sendMessage(msg.chat.id, "🚫 Access denied. You are not authorized to use this bot.");
        return;
    }
    addingMode = true;
    command = 'out';
});

bot.onText(/^\/undo(@[\w_]+)?$/, async (msg) => {
    const chatId = msg.chat.id;
    if (!isUserAllowed(msg.from.id)) {
        bot.sendMessage(chatId, "🚫 Access denied. You are not authorized to use this bot.");
        return;
    }

    const title = msg.chat.title
        ? msg.chat.title.replace(/[^a-zA-Z0-9_-]/g, "").toLowerCase()
        : `user_${msg.chat.id}`;
    const key = `${title}.xlsx`;

    try {
        const workbookBuffer = await downloadExcelBuffer(key);
        let workBook = xlsx.read(workbookBuffer, { type: "buffer" });
        let workSheet = workBook.Sheets["Sheet1"];
        let data = xlsx.utils.sheet_to_json(workSheet);
        if (data.length < 1) {
            bot.sendMessage(chatId, "❌ Table is empty.");
            return;
        }
        data.pop();
        workSheet = xlsx.utils.json_to_sheet(data);
        workBook.Sheets["Sheet1"] = workSheet;

        const newWorkbookBuffer = xlsx.write(workBook, { type: "buffer", bookType: "xlsx" });
        await uploadBuffer(key, newWorkbookBuffer);
        bot.sendMessage(chatId, `✅ Latest entry deleted from ${title}`);
    } catch (error) {
        console.error("Error in /undo:", error);
        bot.sendMessage(chatId, "❌ Table doesn't exist.");
    }
});

const { Readable } = require('stream');

bot.onText(/^\/list(@[\w_]+)?$/, async (msg) => {
    const chatId = msg.chat.id;
    const title = msg.chat.title
        ? msg.chat.title.replace(/[^a-zA-Z0-9_-]/g, "").toLowerCase()
        : `user_${msg.chat.id}`;
    const key = `${title}.xlsx`;
    try {
        const workbookBuffer = await downloadExcelBuffer(key);
        if (!workbookBuffer || workbookBuffer.length === 0) {
            throw new Error("Workbook buffer is empty");
        }
        const readable = new Readable();
        readable.push(workbookBuffer);
        readable.push(null);
        readable.path = `${title}.xlsx`;

        await bot.sendDocument(
            chatId,
            readable,
            { caption: `✅ Latest version of the table ${title}` }
        );
    } catch (error) {
        console.error("Error in /list:", error);
        bot.sendMessage(chatId, `❌ No table found for ${title}.`);
    }
});

bot.on('message', async (msg) => {
    const chatId = msg.chat.id;
    if (!isUserAllowed(msg.from.id) || (msg.text && msg.text.startsWith('/'))) {
        return;
    }
    if (!addingMode) {
        return;
    }

    const parts = msg.text.trim().split(/\s+/);
    let description, amount, percent, inPercentAmount;

    if (command === 'in') {
        if (parts.length !== 3) {
            bot.sendMessage(chatId, "❌ Invalid format. Expected: {description} {amount} {percent}");
            addingMode = false;
            command = '';
            return;
        }
        description = parts[0];
        if (description.trim() === "") {
            bot.sendMessage(chatId, "❌ Invalid description. It should be a non-empty string.");
            addingMode = false;
            command = '';
            return;
        }
        amount = parseInt(parts[1]);
        percent = parseInt(parts[2]);
        if (isNaN(amount) || isNaN(percent)) {
            bot.sendMessage(chatId, "❌ Amount and percent must be valid numbers.");
            addingMode = false;
            command = '';
            return;
        }
        inPercentAmount = percent > 0 ? amount + (amount * percent / 100) : amount - (amount * Math.abs(percent) / 100);
    } else if (command === 'out') {
        if (parts.length !== 2) {
            bot.sendMessage(chatId, "❌ Invalid format. Expected: {description} {amount}");
            addingMode = false;
            command = '';
            return;
        }
        description = parts[0];
        if (description.trim() === "") {
            bot.sendMessage(chatId, "❌ Invalid description. It should be a non-empty string.");
            addingMode = false;
            command = '';
            return;
        }
        amount = parseInt(parts[1]);
        if (isNaN(amount)) {
            bot.sendMessage(chatId, "❌ Amount must be a valid number.");
            addingMode = false;
            command = '';
            return;
        }
    }

    try {
        const title = msg.chat.title
            ? msg.chat.title.replace(/[^a-zA-Z0-9_-]/g, "").toLowerCase()
            : `user_${msg.chat.id}`;
        const key = `${title}.xlsx`;
        let workBook;

        try {
            const workbookBuffer = await downloadExcelBuffer(key);
            workBook = xlsx.read(workbookBuffer, { type: "buffer" });
        } catch (error) {
            workBook = xlsx.utils.book_new();
            const initialSheet = xlsx.utils.aoa_to_sheet([["Date", "Name", "In", "Percent", "Amount with %", "Out", "Balance"]]);
            xlsx.utils.book_append_sheet(workBook, initialSheet, "Sheet1");
        }

        let data = [];
        let workSheet = workBook.Sheets["Sheet1"];
        if (workSheet) {
            data = xlsx.utils.sheet_to_json(workSheet);
        } else {
            workSheet = xlsx.utils.aoa_to_sheet([["Date", "Name", "In", "Percent", "Amount with %", "Out", "Balance"]]);
            workBook.Sheets["Sheet1"] = workSheet;
            xlsx.utils.book_append_sheet(workBook, workSheet, "Sheet1");
        }

        function formatDate(date) {
            const day = String(date.getDate()).padStart(2, "0");
            const month = String(date.getMonth() + 1).padStart(2, "0");
            const year = String(date.getFullYear()).slice(2);
            return `${day}.${month}.${year}`;
        }

        const lastBalance = data.length > 0 ? Number(data[data.length - 1].Balance) || 0 : 0;
        let newBalance = 0;
        if (command === "in") {
            newBalance = lastBalance + inPercentAmount;
        } else if (command === "out") {
            newBalance = lastBalance - amount;
        }

        data.push({
            Date: formatDate(new Date()),
            Name: description,
            In: command === "in" ? amount : "",
            Percent: command === "in" ? percent : "",
            "Amount with %": command === "in" ? inPercentAmount : "",
            Out: command === "out" ? amount : "",
            Balance: newBalance
        });

        workSheet = xlsx.utils.json_to_sheet(data);
        workBook.Sheets["Sheet1"] = workSheet;

        const newWorkbookBuffer = xlsx.write(workBook, { type: "buffer", bookType: "xlsx" });
        await uploadBuffer(key, newWorkbookBuffer);
        bot.sendMessage(chatId, `✅ Entry added to ${title}`);

        addingMode = false;
        command = '';
    } catch (e) {
        console.error("❌ Error:", e);
        bot.sendMessage(chatId, "❌ Failed to update/create table. Please try again.");
    }
});
