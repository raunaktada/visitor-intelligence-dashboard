import dotenv from 'dotenv';
dotenv.config();

import { findSiteAndFileIds } from './services/sharepoint.service';

const siteName = process.argv[2] || '';
const fileName = process.argv[3] || '';

if (!siteName || !fileName) {
  console.error('Usage: ts-node src/find-ids.ts <SiteName> <FileName.xlsx>');
  process.exit(1);
}

findSiteAndFileIds(siteName, fileName).catch(console.error);
