import { NextRequest, NextResponse } from 'next/server';
import * as ExcelJS from 'exceljs';

// Helper function to safely get cell value as string (similar to StatementForm.tsx)
function getCellValueAsString(cell: ExcelJS.Cell | undefined | null): string {
  if (!cell || cell.value === null || cell.value === undefined) {
    return '';
  }
  if (typeof cell.value === 'object' && 'richText' in cell.value) {
    return cell.value.richText.map(rt => rt.text).join('');
  }
  if (typeof cell.value === 'object' && 'result' in cell.value && cell.value.result !== undefined) {
     // For formulas, convert the result to string
    return String(cell.value.result);
  }
   if (cell.value instanceof Date) {
    // Format dates as YYYY-MM-DD or any other preferred format
    // ExcelJS might return them as Date objects if the cell format is date
    return cell.value.toLocaleDateString('en-CA'); // Example: YYYY-MM-DD
  }
  return String(cell.value);
}

// Helper function to get cell value as number (similar to StatementForm.tsx)
function getCellValueAsNumber(cell: ExcelJS.Cell | undefined | null): number {
  if (!cell || cell.value === null || cell.value === undefined) {
    return 0;
  }
  if (typeof cell.value === 'object' && cell.value !== null) {
    if ('result' in cell.value && cell.value.result !== undefined) {
      const num = Number(cell.value.result);
      return isNaN(num) ? 0 : num;
    }
     // If it's a Date object, ExcelJS might have parsed it.
     // We are expecting numbers for Debit/Credit, not date serial numbers here.
    if (cell.value instanceof Date) {
        return 0; 
    }
  }
  const num = Number(cell.value);
  return isNaN(num) ? 0 : num;
}


type StatementEntry = {
    Date: string;
    Particulars: string;
    Debit: number;
    Credit: number;
    Balance?: number; // Optional, as it's not used in summary rows
    Details: string;
};

export async function POST(request: NextRequest) {
    try {
        const formData = await request.formData();
        const file = formData.get('processedStatementFile') as File | null;

        if (!file) {
            return NextResponse.json({ error: 'No file uploaded' }, { status: 400 });
        }

        if (file.type !== 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
            return NextResponse.json({ error: 'Invalid file type. Please upload an .xlsx file.' }, { status: 400 });
        }
        
        const buffer = await file.arrayBuffer();
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);

        const worksheet = workbook.getWorksheet(1); // Assuming data is in the first sheet
        if (!worksheet) {
            return NextResponse.json({ error: 'No worksheet found in the file' }, { status: 400 });
        }

        const entries: StatementEntry[] = [];
        
        // Validate header row: Check if it has at least 6 columns for the expected sequence
        const headerRow = worksheet.getRow(1);
        if (!headerRow || headerRow.cellCount < 6) {
            return NextResponse.json({ error: 'Invalid Excel format: Header row must contain at least 6 columns (Date, Particulars, Debit, Credit, Balance, Details).' }, { status: 400 });
        }

        // Define column indices based on the fixed sequence
        const dateCol = 1;
        const particularsCol = 2;
        const debitCol = 3;
        const creditCol = 4;
        // const balanceCol = 5; // Reserved for Balance, though not directly used in summary items
        const detailsCol = 6;
        
        // Start from row 2 for actual entries
        for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
            const currentRow = worksheet.getRow(rowNumber);
            
            // Skip empty rows or rows that might not have enough cells
            if (currentRow.actualCellCount === 0) continue;

            const details = getCellValueAsString(currentRow.getCell(detailsCol));

            // Skip rows where 'Details' (from the 6th column) is empty or undefined
            if (!details) continue;

            entries.push({
                Date: getCellValueAsString(currentRow.getCell(dateCol)),
                Particulars: getCellValueAsString(currentRow.getCell(particularsCol)),
                Debit: getCellValueAsNumber(currentRow.getCell(debitCol)),
                Credit: getCellValueAsNumber(currentRow.getCell(creditCol)),
                // Balance: getCellValueAsNumber(currentRow.getCell(balanceCol)), // If needed in StatementEntry
                Details: details,
            });
        }

        // Group entries by "Details"
        const groupedEntries: { [key: string]: StatementEntry[] } = entries.reduce((acc, entry) => {
            const groupKey = entry.Details;
            if (!acc[groupKey]) {
                acc[groupKey] = [];
            }
            acc[groupKey].push(entry);
            return acc;
        }, {} as { [key: string]: StatementEntry[] });

        // Create a new workbook for the summary
        const summaryWorkbook = new ExcelJS.Workbook();
        const summarySheet = summaryWorkbook.addWorksheet('Summary');
        
        summarySheet.getColumn('A').width = 15; // Date
        summarySheet.getColumn('B').width = 50; // Particulars
        summarySheet.getColumn('C').width = 15; // Debit
        summarySheet.getColumn('D').width = 15; // Credit


        Object.keys(groupedEntries).forEach(detailName => {
            // Add <Detail name> header
            const detailHeaderRow = summarySheet.addRow([detailName]);
            detailHeaderRow.font = { bold: true, size: 14 };
            summarySheet.mergeCells(detailHeaderRow.number, 1, detailHeaderRow.number, 4); // Merge across 4 columns
            detailHeaderRow.getCell(1).alignment = { horizontal: 'center' };
            summarySheet.addRow([]); // Add a blank row for spacing

            // Add entry headers
            const entryHeaderRow = summarySheet.addRow(['Date', 'Particulars', 'Debit', 'Credit']);
            entryHeaderRow.font = { bold: true };
            entryHeaderRow.getCell('C').alignment = { horizontal: 'right' };
            entryHeaderRow.getCell('D').alignment = { horizontal: 'right' };


            let groupDebitTotal = 0;
            let groupCreditTotal = 0;

            groupedEntries[detailName].forEach(entry => {
                const entryRow = summarySheet.addRow([entry.Date, entry.Particulars, entry.Debit, entry.Credit]);
                entryRow.getCell(3).numFmt = '#,##0.00;[Red]-#,##0.00;0'; // Formatting for Debit
                entryRow.getCell(4).numFmt = '#,##0.00;[Red]-#,##0.00;0'; // Formatting for Credit
                entryRow.getCell(3).alignment = { horizontal: 'right' };
                entryRow.getCell(4).alignment = { horizontal: 'right' };
                groupDebitTotal += entry.Debit;
                groupCreditTotal += entry.Credit;
            });

            // Add <Total> row
            summarySheet.addRow([]); // spacing before total
            const totalRow = summarySheet.addRow(['', 'Total', groupDebitTotal, groupCreditTotal]);
            totalRow.font = { bold: true };
            totalRow.getCell(2).alignment = { horizontal: 'right' };
            totalRow.getCell(3).alignment = { horizontal: 'right' };
            totalRow.getCell(3).numFmt = '#,##0.00;[Red]-#,##0.00;0';
            totalRow.getCell(4).numFmt = '#,##0.00;[Red]-#,##0.00;0';
            
            summarySheet.addRow([]); // Add a blank row for spacing between groups
            summarySheet.addRow([]); // Add another blank row for more spacing
        });

        const outputBuffer = await summaryWorkbook.xlsx.writeBuffer();

        return new NextResponse(outputBuffer, {
            status: 200,
            headers: {
                'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'Content-Disposition': 'attachment; filename="summary_statement.xlsx"',
            },
        });

    } catch (error) {
        console.error('Error processing summary:', error);
        const errorMessage = error instanceof Error ? error.message : 'Failed to generate summary';
        return NextResponse.json({ error: `Server error: ${errorMessage}` }, { status: 500 });
    }
}
