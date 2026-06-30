import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';
import * as dotenv from 'dotenv';

dotenv.config();

async function findFileId() {
  const credential = new ClientSecretCredential(
    process.env.AZURE_TENANT_ID!,
    process.env.AZURE_CLIENT_ID!,
    process.env.AZURE_CLIENT_SECRET!,
  );

  const client = Client.initWithMiddleware({
    authProvider: {
      getAccessToken: async () => {
        const token = await credential.getToken('https://graph.microsoft.com/.default');
        return token!.token;
      },
    },
  });

  try {
    const response = await client
      .api(`/me/drive/root/children?$filter=name eq 'Website Visitors - Sharable Version'`)
      .get();

    if (response.value.length === 0) {
      console.log('File not found. Listing all files in your OneDrive:');
      const allFiles = await client.api('/me/drive/root/children').get();
      allFiles.value.forEach((file: { name: string; id: string }) => {
        console.log(`- ${file.name} (ID: ${file.id})`);
      });
      return;
    }

    const file = response.value[0];
    console.log('✅ File found!');
    console.log(`File name: ${file.name}`);
    console.log(`File ID:   ${file.id}`);
    console.log('\nAdd this to your .env file:');
    console.log(`SHAREPOINT_FILE_ID=${file.id}`);

    console.log('\n--- Worksheets and Tables ---');
    const worksheets = await client
      .api(`/me/drive/items/${file.id}/workbook/worksheets`)
      .get();

    for (const sheet of worksheets.value) {
      console.log(`\nWorksheet: ${sheet.name}`);
      try {
        const tables = await client
          .api(`/me/drive/items/${file.id}/workbook/worksheets('${sheet.name}')/tables`)
          .get();
        tables.value.forEach((t: { name: string }) => console.log(`  - Table: ${t.name}`));
      } catch {
        console.log('  (No tables found)');
      }
    }
  } catch (err) {
    console.error('Error:', err);
  }
}

findFileId();
