require('dotenv').config();
const { MongoClient } = require('mongodb');
const ExcelJS = require('exceljs');

async function main() {
    // Use environment variables for DB connection
    const uri = process.env.MONGODB_URI;
    const dbName = process.env.DB_NAME;
    const collectionName = process.env.DB_COLLECTION || 'users';

    const client = new MongoClient(uri, { useNewUrlParser: true, useUnifiedTopology: true });

    try {
        await client.connect();
        console.log('Connected to MongoDB');
        const db = client.db(dbName);
        const collection = db.collection(collectionName);

        // Fetch all users
        const users = await collection.find({}).toArray();
        console.log(`Fetched ${users.length} users from database`);

        // Sort alphabetically by last name
        users.sort((a, b) => {
            if (!a.lastName) return 1;
            if (!b.lastName) return -1;
            return a.lastName.localeCompare(b.lastName);
        });

        // Prepare data for Excel: only firstName, lastName, email
        const exportData = users.map(user => ({
            firstName: user.firstName || '',
            lastName: user.lastName || '',
            email: user.email || ''
        }));

        // Create Excel workbook and worksheet
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Users');
        worksheet.columns = [
            { header: 'First Name', key: 'firstName', width: 20 },
            { header: 'Last Name', key: 'lastName', width: 20 },
            { header: 'Email', key: 'email', width: 30 }
        ];

        worksheet.addRows(exportData);

        // Save Excel file in the script directory
        const fileName = 'users_export.xlsx';
        await workbook.xlsx.writeFile(fileName);
        console.log(`Exported data to ${fileName}`);
    } catch (e) {
        console.error(e);
        process.exit(1);
    } finally {
        await client.close();
    }
}

main();
