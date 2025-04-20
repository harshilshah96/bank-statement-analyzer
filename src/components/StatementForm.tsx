'use client';

import React, { ChangeEvent, useState } from 'react';
import { useForm, useFieldArray, SubmitHandler } from 'react-hook-form';

type KeywordPair = { keyword: string; value: string };

type FormValues = {
  keywordsFile: FileList | null;
  bankStatementFile: FileList | null;
  keywords: KeywordPair[];
};

export default function StatementForm() {
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [submitStatus, setSubmitStatus] = useState<'success' | 'error' | null>(null);
  const [submitMessage, setSubmitMessage] = useState<string>('');

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

  const onSubmit: SubmitHandler<FormValues> = async (data) => {
     console.log("Submitting final form data:", data);
     setIsSubmitting(true);
     setSubmitStatus(null);
     setSubmitMessage('');

     if (!data.bankStatementFile || data.bankStatementFile.length === 0) {
       alert('Please upload the bank statement file.');
       setIsSubmitting(false);
       return;
     }

     const finalFormData = new FormData();

     finalFormData.append('bankStatementFile', data.bankStatementFile[0]);
     finalFormData.append('keywords', JSON.stringify(data.keywords));

     try {
       const response = await fetch('/api/submit', {
         method: 'POST',
         body: finalFormData,
       });

       if (!response.ok) {
         let errorMsg = `HTTP error! status: ${response.status} ${response.statusText}`;
         try {
             const errData = await response.json();
             errorMsg = errData.error || errorMsg;
         } catch (e) { 
          console.error('Error submitting main form:', e);
          }
         throw new Error(errorMsg);
       }

       const contentType = response.headers.get('content-type');
       if (!contentType || !contentType.includes('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')) {
           throw new Error('Received unexpected response format from server.');
       }

       const blob = await response.blob();

       const url = window.URL.createObjectURL(blob);
       const a = document.createElement('a');
       a.style.display = 'none';
       a.href = url;
       const disposition = response.headers.get('content-disposition');
       let filename = 'processed_statement.xlsx';
       if (disposition && disposition.indexOf('attachment') !== -1) {
           const filenameRegex = /filename[^;=\n]*=\s*("([^"]*)"|([^;\n]*))/;
           const matches = filenameRegex.exec(disposition);
           if (matches != null && (matches[2] || matches[3])) {
               filename = matches[2] || matches[3];
           }
       }
       a.download = filename;
       document.body.appendChild(a);
       a.click();
       window.URL.revokeObjectURL(url);
       document.body.removeChild(a);

       setSubmitStatus('success');
       setSubmitMessage('Processing complete! Your download should start shortly.');

     } catch (error) {
       console.error('Error submitting main form:', error);
       setSubmitStatus('error');
       setSubmitMessage(`Submission failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
     } finally {
       setIsSubmitting(false);
     }
  };

  return (
    <form onSubmit={handleSubmit(onSubmit)} className="space-y-6 p-4 max-w-lg mx-auto bg-white shadow-md rounded-lg">
      <div>
        <label htmlFor="keywordsFile" className="block text-sm font-medium text-gray-700 mb-1">
          1. Upload Keywords (Excel or CSV)
        </label>
        <input
          id="keywordsFile"
          type="file"
          accept=".xlsx,.xls,.csv"
          {...register('keywordsFile', {
             onChange: handleKeywordsFileChange
            })}
          className="block w-full text-sm
                     file:mr-4 file:py-2 file:px-4
                     file:rounded-full file:border-0
                     file:text-sm file:font-semibold
                     file:bg-blue-50 file:text-blue-700
                     hover:file:bg-blue-100"
        />
        {errors.keywordsFile && <p className="mt-1 text-sm text-red-600">{errors.keywordsFile.message}</p>}
      </div>

      <div>
        <h3 className="text-lg font-medium text-gray-900 mb-2">Keywords (from file)</h3>
        {fields.map((item, index) => (
          <div key={item.id} className="flex space-x-2 mb-2 items-center">
            <div className="flex-1">
              <label htmlFor={`keywords.${index}.keyword`} className="sr-only">Keyword</label>
              <input
                id={`keywords.${index}.keyword`}
                {...register(`keywords.${index}.keyword`, { required: 'Keyword is required' })}
                placeholder="Keyword"
                className="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-indigo-500 focus:border-indigo-500 sm:text-sm text-black bg-gray-50"
                readOnly
              />
              {errors.keywords?.[index]?.keyword && <p className="mt-1 text-sm text-red-600">{errors.keywords[index]?.keyword?.message}</p>}
            </div>
            <div className="flex-1">
              <label htmlFor={`keywords.${index}.value`} className="sr-only">Value</label>
              <input
                id={`keywords.${index}.value`}
                {...register(`keywords.${index}.value`, { required: 'Value is required' })}
                placeholder="Value"
                className="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-indigo-500 focus:border-indigo-500 sm:text-sm text-black bg-gray-50"
                readOnly
              />
              {errors.keywords?.[index]?.value && <p className="mt-1 text-sm text-red-600">{errors.keywords[index]?.value?.message}</p>}
            </div>
          </div>
        ))}
      </div>

      <div>
        <label htmlFor="bankStatementFile" className="block text-sm font-medium text-gray-700 mb-1">
          2. Upload Bank Statement (Excel or CSV)
        </label>
        <input
          id="bankStatementFile"
          type="file"
          accept=".xlsx,.xls,.csv"
          {...register('bankStatementFile', { required: 'Bank statement file is required' })}
          className="block w-full text-sm
                     file:mr-4 file:py-2 file:px-4
                     file:rounded-full file:border-0
                     file:text-sm file:font-semibold
                     file:bg-green-50 file:text-green-700
                     hover:file:bg-green-100"
          disabled={fields.length <= 1 && !fields[0]?.keyword}
        />
        {errors.bankStatementFile && <p className="mt-1 text-sm text-red-600">{errors.bankStatementFile.message}</p>}
        {fields.length <= 1 && !fields[0]?.keyword && <p className="mt-1 text-sm text-gray-500">Upload keywords file first.</p>}
      </div>

      <div>
        <button
          type="submit"
          className={`w-full flex justify-center py-2 px-4 border border-transparent rounded-md shadow-sm text-sm font-medium text-white ${isSubmitting ? 'bg-gray-400' : 'bg-indigo-600 hover:bg-indigo-700'} focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500`}
          disabled={isSubmitting}
        >
          {isSubmitting ? 'Processing...' : 'Process Statement & Download'}
        </button>
        {submitStatus === 'success' && <p className="mt-2 text-sm text-green-600">{submitMessage}</p>}
        {submitStatus === 'error' && <p className="mt-2 text-sm text-red-600">{submitMessage}</p>}
      </div>
    </form>
  );
} 