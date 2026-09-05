const { PrismaClient } = require('@prisma/client');
const fs = require('fs');
const path = require('path');

const prisma = new PrismaClient();
const outDir = path.join(__dirname, '../SQL_Reference');

async function main() {
  console.log("Connecting to SQL Server via Prisma to extract legacy logic...");
  
  if (!fs.existsSync(outDir)) {
    fs.mkdirSync(outDir, { recursive: true });
  }

  try {
    const results = await prisma.$queryRaw`
      SELECT 
        o.name AS object_name,
        o.type_desc AS object_type,
        sm.definition AS sql_code
      FROM sys.sql_modules sm
      JOIN sys.objects o ON sm.object_id = o.object_id
      WHERE o.type IN ('P', 'V', 'TR') -- P = Stored Proc, V = View, TR = Trigger
        AND o.is_ms_shipped = 0 -- Exclude system objects
    `;

    console.log(`Found ${results.length} legacy SQL objects.`);

    for (const obj of results) {
      let ext = '.sql';
      let typeFolder = 'Unknown';

      if (obj.object_type === 'SQL_STORED_PROCEDURE') typeFolder = 'StoredProcedures';
      else if (obj.object_type === 'VIEW') typeFolder = 'Views';
      else if (obj.object_type === 'SQL_TRIGGER') typeFolder = 'Triggers';

      const typeDir = path.join(outDir, typeFolder);
      if (!fs.existsSync(typeDir)) {
        fs.mkdirSync(typeDir, { recursive: true });
      }

      const safeName = obj.object_name.replace(/[^a-zA-Z0-9_]/g, '_');
      const filePath = path.join(typeDir, `${safeName}${ext}`);
      const sqlCode = obj.sql_code || `-- [No Definition Found / Encrypted Object]\n-- Object Type: ${obj.object_type}`;
      
      fs.writeFileSync(filePath, sqlCode);
    }
    
    console.log(`Successfully extracted ${results.length} SQL objects to Migration_Ref/SQL_Reference/`);
  } catch (error) {
    console.error("Failed to extract SQL logic. Please ensure Prisma is initialized and the database is reachable.");
    console.error(error);
  } finally {
    await prisma.$disconnect();
  }
}

main();
