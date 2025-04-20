import { NextRequest, NextResponse } from 'next/server';
import { openai } from '@ai-sdk/openai';
import { generateText } from 'ai';

// Ensure OPENAI_API_KEY is loaded (Vercel AI SDK reads it automatically)
// You might need to configure environment variables differently depending on deployment

type KeywordPair = {
    keyword: string;
    value: string;
};

type RequestData = {
    particulars: string;
    debit: number;
    credit: number;
    keywords: KeywordPair[];
};

export async function POST(request: NextRequest) {
    try {
        const data: RequestData = await request.json();
        
        // Validate required fields
        if (data.particulars === undefined) {
            return NextResponse.json({ error: 'Particulars field is missing' }, { status: 400 });
        }
        if (data.debit === undefined) {
            return NextResponse.json({ error: 'Debit field is missing' }, { status: 400 });
        }
        if (data.credit === undefined) {
            return NextResponse.json({ error: 'Credit field is missing' }, { status: 400 });
        }
        if (!Array.isArray(data.keywords)) {
            return NextResponse.json({ error: 'Keywords must be an array' }, { status: 400 });
        }

        let details = '';

        // Only process if there is actual transaction data
        if (!data.particulars && data.debit === 0 && data.credit === 0) {
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
- Particulars (Comment): ${data.particulars}
- Debit amount: ${data.debit}
- Credit amount: ${data.credit}

Keyword List:
${data.keywords.map(kw => `- ${kw.keyword}: ${kw.value}`).join('\n')}

Determine the 'Details' category based *only* on the rules provided.`;

                const { text } = await generateText({
                    // Pass the provider and model ID directly
                    model: openai('gpt-4o-mini'),
                    system: systemPrompt,
                    prompt: userPrompt,
                });
                
                details = text.trim();
                console.log(`Processing details: ${details}`);

            } catch (aiError) {
                console.error(`AI processing error:`, aiError);
                return NextResponse.json({ error: 'AI processing failed' }, { status: 500 });
            }
        }

        return NextResponse.json({ details });

    } catch (error) {
        console.error('Error processing request:', error);
        // Provide a more specific error message if possible
        const errorMessage = error instanceof Error ? error.message : 'Failed to process request';
        return NextResponse.json({ error: `Server error: ${errorMessage}` }, { status: 500 });
    }
} 