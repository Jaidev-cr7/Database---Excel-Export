require('dotenv').config();
const express = require('express');
const { MongoClient } = require('mongodb');
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 3000;

app.get('/export-users', async (req, res) => {
    const uri = process.env.MONGODB_URI;
    const dbName = process.env.DB_NAME;
    const collectionName = process.env.DB_COLLECTION || 'users';
    const client = new MongoClient(uri, { useNewUrlParser: true, useUnifiedTopology: true });

    try {
        await client.connect();
        const db = client.db(dbName);
        const collection = db.collection(collectionName);
        const users = await collection.find({}).toArray();
        users.sort((a, b) => {
            if (!a.lastName) return 1;
            if (!b.lastName) return -1;
            return a.lastName.localeCompare(b.lastName);
        });
        const exportData = users.map(user => ({
            firstName: user.firstName || '',
            lastName: user.lastName || '',
            email: user.email || ''
        }));
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Users');
        worksheet.columns = [
            { header: 'First Name', key: 'firstName', width: 20 },
            { header: 'Last Name', key: 'lastName', width: 20 },
            { header: 'Email', key: 'email', width: 30 }
        ];
        worksheet.addRows(exportData);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=users_export.xlsx');
        await workbook.xlsx.write(res);
        res.end();
    } catch (e) {
        console.error(e);
        res.status(500).send('Error exporting users');
    } finally {
        await client.close();
    }
});

app.get('/', (req, res) => {
    res.send('User Export API is running. Use /export-users to download Excel.');
});

app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
