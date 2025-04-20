import { NextRequest, NextResponse } from 'next/server';
import * as ExcelJS from 'exceljs';
import { Readable } from 'stream';
import { openai } from '@ai-sdk/openai';
import { generateText } from 'ai';

// Ensure OPENAI_API_KEY is loaded (Vercel AI SDK reads it automatically)
// You might need to configure environment variables differently depending on deployment

type KeywordPair = {
    keyword: string;
    value: string;
};

// Helper function to safely get cell value as string
function getCellValueAsString(cell: ExcelJS.Cell | undefined | null): string {
    if (!cell || cell.value === null || cell.value === undefined) {
        return '';
    }
    if (typeof cell.value === 'object' && 'richText' in cell.value) {
        // Handle RichTextValue
        return cell.value.richText.map(rt => rt.text).join('');
    }
    if (typeof cell.value === 'object' && 'result' in cell.value) {
        // Handle FormulaValue - use the calculated result
        return String(cell.value.result || '');
    }
    return String(cell.value);
}

// Helper function to get cell value as number
function getCellValueAsNumber(cell: ExcelJS.Cell | undefined | null): number {
    if (!cell || cell.value === null || cell.value === undefined) {
        return 0;
    }
    const num = Number(cell.value);
    return isNaN(num) ? 0 : num;
}


export async function POST(request: NextRequest) {
    try {
        const formData = await request.formData();

        const bankStatementFile = formData.get('bankStatementFile') as File | null;
        const keywordsString = formData.get('keywords') as string | null;

        if (!bankStatementFile) {
            return NextResponse.json({ error: 'Bank statement file is missing' }, { status: 400 });
        }
        if (!keywordsString) {
            return NextResponse.json({ error: 'Keywords data is missing' }, { status: 400 });
        }

        let keywords: KeywordPair[];
        try {
            keywords = JSON.parse(keywordsString);
            if (!Array.isArray(keywords)) throw new Error('Keywords is not an array');
        } catch (e) {
            console.error('Error parsing keywords:', e);
            return NextResponse.json({ error: 'Invalid keywords format' }, { status: 400 });
        }

        // --- Load Bank Statement Workbook ---
        const inputWorkbook = new ExcelJS.Workbook();
        const buffer = await bankStatementFile.arrayBuffer();

        if (bankStatementFile.type === 'text/csv') {
            const stream = Readable.from(Buffer.from(buffer));
            await inputWorkbook.csv.read(stream);
        } else if (['application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'].includes(bankStatementFile.type)){
            await inputWorkbook.xlsx.load(buffer);
        } else {
             return NextResponse.json({ error: 'Invalid bank statement file type.' }, { status: 400 });
        }

        const inputWorksheet = inputWorkbook.getWorksheet(1);
        if (!inputWorksheet) {
            return NextResponse.json({ error: 'No worksheet found in the bank statement file' }, { status: 400 });
        }

        // --- Create Output Workbook ---
        const outputWorkbook = new ExcelJS.Workbook();
        const outputWorksheet = outputWorkbook.addWorksheet('Processed Statement');

        // --- Process Headers ---
        const headerRow = inputWorksheet.getRow(1);
        const headers: (string | ExcelJS.RichText)[] = [];
        headerRow.eachCell((cell) => {
            headers.push(cell.value as string | ExcelJS.RichText); // Assume headers are simple strings or RichText
        });
        outputWorksheet.addRow([...headers, 'Details']); // Add the new Details header
        outputWorksheet.getRow(1).font = { bold: true }; // Make header bold

        // --- Process Rows with AI ---
        for (let rowNumber = 2; rowNumber <= inputWorksheet.rowCount; rowNumber++) {
            const inputRow = inputWorksheet.getRow(rowNumber);

            // Assuming columns: Date[1], Particulars[2], Debit[3], Credit[4], Balance[5]
            // Adjust indices if your statement structure is different
            const dateVal = inputRow.getCell(1).value; // Keep original date format if possible
            const particulars = getCellValueAsString(inputRow.getCell(2));
            const debit = getCellValueAsNumber(inputRow.getCell(3));
            const credit = getCellValueAsNumber(inputRow.getCell(4));
            const balanceVal = inputRow.getCell(5).value; // Keep original balance format

            let details = 'Error Processing';

            if (!particulars && debit === 0 && credit === 0) {
                details = '';
            } else {
                 try {
                    const systemPrompt = `You are an assistant that categorizes bank statement transactions based on keywords and transaction type. 
Your task is to determine the 'Details' category for a given transaction.
Follow these rules precisely:
1. Examine the 'Particulars (Comment)' of the transaction.
2. Compare the 'Particulars (Comment)' semantically against the provided 'Keyword':'Value' pairs.
3. If you find a 'Keyword' that is semantically similar or a close match to the 'Particulars (Comment)', return the corresponding 'Value' exactly as provided.
4. If no semantic match is found with any keyword:
   a. Check if 'Debit amount' is greater than 0. If yes, return 'Personal expense'.
   b. If 'Debit amount' is 0 or less, check if 'Credit amount' is greater than 0. If yes, return 'Business income'.
   c. If neither 'Debit amount' nor 'Credit amount' is greater than 0, return 'Uncategorized'.
5. Respond ONLY with the determined category string ('Value' from keyword pair, 'Personal expense', 'Business income', or 'Uncategorized'). Do not add any explanation or introductory text.`;

                    const userPrompt = `Transaction:
- Particulars (Comment): ${particulars}
- Debit amount: ${debit}
- Credit amount: ${credit}

Keyword List:
${keywords.map(kw => `- ${kw.keyword}: ${kw.value}`).join('\n')}

Determine the 'Details' category based *only* on the rules provided.`;

                    const { text } = await generateText({
                        // Pass the provider and model ID directly
                        model: openai('gpt-4o-mini'),
                        system: systemPrompt,
                        prompt: userPrompt,
                    });
                    // Add logging to check the output
                    console.log(`Row ${rowNumber} details: ${text}`);
                    details = text.trim();

                } catch (aiError) {
                    console.error(`AI processing error for row ${rowNumber}:`, aiError);
                }
            }

            // Add row to output sheet
            const originalRowValues: any[] = [];
            inputRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                 // Attempt to preserve original cell value types (date, number, string)
                 // inputRow.getCell(colNumber).value might be simpler if formatting isn't critical
                 if (colNumber === 1 || colNumber === 5) { // Date and Balance
                     originalRowValues.push(inputRow.getCell(colNumber).value);
                 } else if (colNumber === 3 || colNumber === 4) { // Debit/Credit
                     originalRowValues.push(getCellValueAsNumber(inputRow.getCell(colNumber)));
                 } else { // Particulars and any other columns
                      originalRowValues.push(getCellValueAsString(inputRow.getCell(colNumber)));
                 }
            });

            // Ensure we match the number of original headers
            const paddedOriginalValues = originalRowValues.slice(0, headers.length);
            while (paddedOriginalValues.length < headers.length) {
                paddedOriginalValues.push('');
            }

            outputWorksheet.addRow([...paddedOriginalValues, details]);

             // Optional: Add a small delay to avoid hitting rate limits if processing many rows
             // await new Promise(resolve => setTimeout(resolve, 50));
        }

        // --- Return Processed File ---
        const outputBuffer = await outputWorkbook.xlsx.writeBuffer();

        return new NextResponse(outputBuffer, {
            status: 200,
            headers: {
                'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'Content-Disposition': 'attachment; filename="processed_statement.xlsx"',
            },
        });

    } catch (error) {
        console.error('Error processing form:', error);
        // Provide a more specific error message if possible
        const errorMessage = error instanceof Error ? error.message : 'Failed to process request';
        return NextResponse.json({ error: `Server error: ${errorMessage}` }, { status: 500 });
    }
} 