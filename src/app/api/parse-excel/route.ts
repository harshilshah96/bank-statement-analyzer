import { NextRequest, NextResponse } from 'next/server';
import * as ExcelJS from 'exceljs';
import { Readable } from 'stream'; // Import Readable from stream

type KeywordPair = {
  keyword: string;
  value: string;
};

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get('statementFile') as File | null;

    if (!file) {
      return NextResponse.json({ error: 'No file uploaded' }, { status: 400 });
    }

    // Check file type (optional but recommended)
    if (!['application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'text/csv'].includes(file.type)) {
       // Allow CSV as well, adjust mime types if needed
      // return NextResponse.json({ error: 'Invalid file type. Only Excel (.xlsx, .xls) or CSV (.csv) files are allowed.' }, { status: 400 });
    }

    const buffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook(); // Use ExcelJS.default.Workbook() if you encounter constructor issues

    // Determine if it's CSV or Excel and load accordingly
    if (file.type === 'text/csv') {
        // Create a readable stream from the buffer for CSV parsing
        const stream = Readable.from(Buffer.from(buffer));
        await workbook.csv.read(stream);
    } else {
        await workbook.xlsx.load(buffer);
    }


    const worksheet = workbook.getWorksheet(1); // Get the first worksheet
    if (!worksheet) {
        return NextResponse.json({ error: 'No worksheet found in the file' }, { status: 400 });
    }

    const keywords: KeywordPair[] = [];
    // Iterate over rows starting from the second row (index 2) to skip the header
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber > 1) { // Skip header row (rowNumber is 1-based)
        const keywordCell = row.getCell(1); // Column A: Statement
        const valueCell = row.getCell(2);   // Column B: Remarks

        const keyword = keywordCell.value ? String(keywordCell.value) : '';
        const value = valueCell.value ? String(valueCell.value) : '';

        if (keyword || value) { // Add row if at least one cell has content
            keywords.push({ keyword, value });
        }
      }
    });

    return NextResponse.json({ keywords });

  } catch (error) {
    console.error('Error parsing Excel file:', error);
    // Check for specific exceljs errors if needed
    return NextResponse.json({ error: 'Failed to parse Excel file' }, { status: 500 });
  }
} 