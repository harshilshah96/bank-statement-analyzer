'use client';

import React, { ChangeEvent, useState } from 'react';
import { useForm, useFieldArray, SubmitHandler } from 'react-hook-form';
import * as ExcelJS from 'exceljs';

type KeywordPair = { keyword: string; value: string };

type FormValues = {
  keywordsFile: FileList | null;
  bankStatementFile: FileList | null;
  keywords: KeywordPair[];
};

type ProcessedRow = {
  [key: string]: string | number | Date | null;
  details: string;
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

export default function StatementForm() {
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [submitStatus, setSubmitStatus] = useState<'success' | 'error' | 'processing' | null>(null);
  const [submitMessage, setSubmitMessage] = useState<string>('');
  const [processedRows, setProcessedRows] = useState<ProcessedRow[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [processedCount, setProcessedCount] = useState(0);
  const [totalRows, setTotalRows] = useState(0);
  
  const {
    register,
    control,
    handleSubmit,
    formState: { errors },
    reset,
    getValues,
  } = useForm<FormValues>({
    defaultValues: {
      keywordsFile: null,
      bankStatementFile: null,
      keywords: [{ keyword: '', value: '' }],
    },
  });

  const { fields } = useFieldArray({
    control,
    name: 'keywords',
  });

  const handleKeywordsFileChange = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files ? event.target.files[0] : null;
    setSubmitStatus(null);
    setSubmitMessage('');
    if (!file) {
      return;
    }

    const allowedTypes = [
      'application/vnd.ms-excel',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'text/csv'
    ];
    if (!allowedTypes.includes(file.type)) {
        alert('Invalid keywords file type. Please upload an Excel (.xlsx, .xls) or CSV (.csv) file.');
        event.target.value = '';
        reset({ ...getValues(), keywordsFile: null, keywords: [{ keyword: '', value: '' }] });
        return;
    }

    const formData = new FormData();
    formData.append('statementFile', file);

    try {
      const response = await fetch('/api/parse-excel', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || `Failed to parse file: ${response.statusText}`);
      }

      const result = await response.json();

      if (result.keywords && Array.isArray(result.keywords)) {
        const currentValues = getValues();
        reset({
          ...currentValues,
          keywordsFile: event.target.files,
          keywords: result.keywords.length > 0 ? result.keywords : [{ keyword: '', value: '' }],
        });
      } else {
        console.warn('Parsed keywords data is not in the expected format:', result);
        reset({ ...getValues(), keywords: [{ keyword: '', value: '' }] });
      }

    } catch (error) {
      console.error('Error processing keywords file:', error);
      alert(`Error parsing keywords file: ${error instanceof Error ? error.message : 'Could not process file'}`);
      reset({ ...getValues(), keywordsFile: null, keywords: [{ keyword: '', value: '' }] });
      event.target.value = '';
    }
  };

  // Function to process a single row with API call
  const processRow = async (
    particulars: string, 
    debit: number, 
    credit: number, 
    keywords: KeywordPair[]
  ): Promise<string> => {
    try {
      const response = await fetch('/api/submit', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          particulars,
          debit,
          credit,
          keywords,
        }),
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || `API error: ${response.statusText}`);
      }

      const result = await response.json();
      return result.details || 'Uncategorized';
    } catch (error) {
      console.error('Error processing row:', error);
      return 'Error Processing';
    }
  };

  const onSubmit: SubmitHandler<FormValues> = async (data) => {
    console.log("Submitting final form data:", data);
    setIsSubmitting(true);
    setSubmitStatus('processing');
    setSubmitMessage('Processing statement...');
    setProcessedRows([]);
    setHeaders([]);
    setProcessedCount(0);
    setTotalRows(0);

    if (!data.bankStatementFile || data.bankStatementFile.length === 0) {
      alert('Please upload the bank statement file.');
      setIsSubmitting(false);
      setSubmitStatus(null);
      return;
    }

    try {
      // Process the bank statement file on the client side
      const bankStatementFile = data.bankStatementFile[0];
      const inputWorkbook = new ExcelJS.Workbook();
      const buffer = await bankStatementFile.arrayBuffer();

      if (bankStatementFile.type === 'text/csv') {
        // Use string parsing for CSV in browser environment
        throw new Error('CSV files are not supported. Please upload an Excel (.xlsx) file instead.');
      } else {
        await inputWorkbook.xlsx.load(buffer);
      }

      const inputWorksheet = inputWorkbook.getWorksheet(1);
      if (!inputWorksheet) {
        throw new Error('No worksheet found in the bank statement file');
      }

      // Process Headers
      const headerRow = inputWorksheet.getRow(1);
      const headerValues: string[] = [];
      headerRow.eachCell((cell) => {
        headerValues.push(getCellValueAsString(cell));
      });
      headerValues.push('Details'); // Add the new Details header
      
      setHeaders(headerValues);
      setTotalRows(inputWorksheet.rowCount - 1); // Exclude header row
      
      // Create output workbook
      const outputWorkbook = new ExcelJS.Workbook();
      const outputWorksheet = outputWorkbook.addWorksheet('Processed Statement');
      
      // Add headers to output worksheet
      outputWorksheet.addRow(headerValues);
      outputWorksheet.getRow(1).font = { bold: true };

      // Process each row
      const allProcessedRows: ProcessedRow[] = [];
      
      for (let rowNumber = 2; rowNumber <= inputWorksheet.rowCount; rowNumber++) {
        const inputRow = inputWorksheet.getRow(rowNumber);
        
        // Assuming columns: Date[1], Particulars[2], Debit[3], Credit[4], Balance[5]
        // Adjust indices if your statement structure is different
        const particulars = getCellValueAsString(inputRow.getCell(2));
        const debit = getCellValueAsNumber(inputRow.getCell(3));
        const credit = getCellValueAsNumber(inputRow.getCell(4));
        
        // Create an object to store all cell values
        const rowData: ProcessedRow = {
          details: '',
        };
        
        // Add all original cell values to the row data
        inputRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const headerKey = headerValues[colNumber - 1];
          if (headerKey) {
            if (colNumber === 1) { // Date column
              rowData[headerKey] = cell.value as Date;
            } else if (colNumber === 3 || colNumber === 4) { // Debit/Credit columns
              rowData[headerKey] = getCellValueAsNumber(cell);
            } else {
              rowData[headerKey] = getCellValueAsString(cell);
            }
          }
        });
        
        // Process the row with AI if it has content
        if (!particulars && debit === 0 && credit === 0) {
          rowData.details = '';
        } else {
          // Make API call to process this row
          const details = await processRow(particulars, debit, credit, data.keywords);
          rowData.details = details;
        }
        
        // Add to processed rows and update state
        allProcessedRows.push(rowData);
        setProcessedRows([...allProcessedRows]);
        setProcessedCount(rowNumber - 1);
        
        // Add row to output worksheet
        const rowForExcel = headerValues.map((header, index) => {
          if (index === headerValues.length - 1) { // Details column
            return rowData.details;
          } else {
            return rowData[header] ?? '';
          }
        });
        
        outputWorksheet.addRow(rowForExcel);
      }
      
      // Generate Excel file for download
      const outputBuffer = await outputWorkbook.xlsx.writeBuffer();
      const blob = new Blob([outputBuffer], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.style.display = 'none';
      a.href = url;
      a.download = 'processed_statement.xlsx';
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
      
      setSubmitStatus('success');
      setSubmitMessage('Processing complete! Your download should start shortly.');
      
    } catch (error) {
      console.error('Error processing bank statement:', error);
      setSubmitStatus('error');
      setSubmitMessage(`Processing failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <div className="space-y-6 p-4 max-w-4xl mx-auto">
      <form onSubmit={handleSubmit(onSubmit)} className="space-y-6 p-4 bg-gray-800 shadow-lg rounded-lg">
        <div>
          <label htmlFor="keywordsFile" className="block text-sm font-medium text-gray-300 mb-1">
            1. Upload Keywords (Excel or CSV)
          </label>
          <input
            id="keywordsFile"
            type="file"
            accept=".xlsx,.xls,.csv"
            {...register('keywordsFile', {
              onChange: handleKeywordsFileChange
              })}
            className="block w-full text-sm text-gray-300
                      file:mr-4 file:py-2 file:px-4
                      file:rounded-full file:border-0
                      file:text-sm file:font-semibold
                      file:bg-blue-700 file:text-white
                      hover:file:bg-blue-600
                      focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 focus:ring-offset-gray-800"
          />
          {errors.keywordsFile && <p className="mt-1 text-sm text-red-400">{errors.keywordsFile.message}</p>}
        </div>

        <div>
          <h3 className="text-lg font-medium text-gray-100 mb-2">Keywords (from file)</h3>
          <div className="max-h-60 overflow-y-auto pr-2">
            {fields.map((item, index) => (
              <div key={item.id} className="flex space-x-2 mb-2 items-center">
                <div className="flex-1">
                  <label htmlFor={`keywords.${index}.keyword`} className="sr-only">Keyword</label>
                  <input
                    id={`keywords.${index}.keyword`}
                    {...register(`keywords.${index}.keyword`, { required: 'Keyword is required' })}
                    placeholder="Keyword"
                    className="mt-1 block w-full px-3 py-2 border border-gray-600 rounded-md shadow-sm focus:outline-none focus:ring-indigo-500 focus:border-indigo-500 sm:text-sm text-gray-200 bg-gray-700 placeholder-gray-400"
                  />
                  {errors.keywords?.[index]?.keyword && <p className="mt-1 text-sm text-red-400">{errors.keywords[index]?.keyword?.message}</p>}
                </div>
                <div className="flex-1">
                  <label htmlFor={`keywords.${index}.value`} className="sr-only">Value</label>
                  <input
                    id={`keywords.${index}.value`}
                    {...register(`keywords.${index}.value`, { required: 'Value is required' })}
                    placeholder="Value"
                    className="mt-1 block w-full px-3 py-2 border border-gray-600 rounded-md shadow-sm focus:outline-none focus:ring-indigo-500 focus:border-indigo-500 sm:text-sm text-gray-200 bg-gray-700 placeholder-gray-400"
                  />
                  {errors.keywords?.[index]?.value && <p className="mt-1 text-sm text-red-400">{errors.keywords[index]?.value?.message}</p>}
                </div>
              </div>
            ))}
          </div>
        </div>

        <div>
          <label htmlFor="bankStatementFile" className="block text-sm font-medium text-gray-300 mb-1">
            2. Upload Bank Statement (Excel or CSV)
          </label>
          <input
            id="bankStatementFile"
            type="file"
            accept=".xlsx,.xls,.csv"
            {...register('bankStatementFile', { required: 'Bank statement file is required' })}
            className="block w-full text-sm text-gray-300
                      file:mr-4 file:py-2 file:px-4
                      file:rounded-full file:border-0
                      file:text-sm file:font-semibold
                      file:bg-green-700 file:text-white
                      hover:file:bg-green-600
                      focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-green-500 focus:ring-offset-gray-800
                      disabled:opacity-50 disabled:cursor-not-allowed"
            disabled={fields.length <= 1 && !fields[0]?.keyword}
          />
          {errors.bankStatementFile && <p className="mt-1 text-sm text-red-400">{errors.bankStatementFile.message}</p>}
          {fields.length <= 1 && !fields[0]?.keyword && <p className="mt-1 text-sm text-gray-500">Upload keywords file first.</p>}
        </div>

        <div>
          <button
            type="submit"
            className={`w-full flex justify-center py-2 px-4 border border-transparent rounded-md shadow-sm text-sm font-medium text-white ${isSubmitting ? 'bg-gray-600 cursor-not-allowed' : 'bg-indigo-600 hover:bg-indigo-700'} focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 focus:ring-offset-gray-800`}
            disabled={isSubmitting}
          >
            {isSubmitting ? 'Processing...' : 'Process Statement & Download'}
          </button>
          {submitStatus === 'success' && <p className="mt-2 text-sm text-green-400">{submitMessage}</p>}
          {submitStatus === 'error' && <p className="mt-2 text-sm text-red-400">{submitMessage}</p>}
          {submitStatus === 'processing' && (
            <div className="mt-2">
              <p className="text-sm text-blue-400">{submitMessage}</p>
              <p className="text-xs text-gray-400">Processing row {processedCount} of {totalRows}</p>
              <div className="w-full bg-gray-700 rounded-full h-2.5 mt-1">
                <div
                  className="bg-blue-600 h-2.5 rounded-full"
                  style={{ width: `${totalRows ? (processedCount / totalRows) * 100 : 0}%` }}
                ></div>
              </div>
            </div>
          )}
        </div>
      </form>

      {/* Results Table */}
      {processedRows.length > 0 && (
        <div className="mt-6 p-4 bg-gray-800 shadow-lg rounded-lg overflow-x-auto">
          <h3 className="text-lg font-medium text-gray-100 mb-2">Processing Results</h3>
          <div className="max-h-[400px] overflow-y-auto relative">
            <table className="min-w-full divide-y divide-gray-700">
              <thead className="bg-gray-700 sticky top-0 z-10">
                <tr>
                  {headers.map((header, index) => (
                    <th 
                      key={index} 
                      className="px-6 py-3 text-left text-xs font-medium text-gray-300 uppercase tracking-wider"
                    >
                      {header}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="bg-gray-800 divide-y divide-gray-700">
                {processedRows.map((row, rowIndex) => (
                  <tr key={rowIndex} className={rowIndex % 2 === 0 ? 'bg-gray-800' : 'bg-gray-750'}>
                    {headers.map((header, colIndex) => (
                      <td 
                        key={`${rowIndex}-${colIndex}`} 
                        className="px-6 py-4 whitespace-nowrap text-sm text-gray-300"
                      >
                        {header === 'Details' 
                          ? row.details 
                          : row[header] !== null && row[header] !== undefined 
                            ? String(row[header]) 
                            : ''}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
} 