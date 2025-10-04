import { NextRequest } from 'next/server';
import { unlockExcelBuffer, isLikelyExcelZip } from '@/lib/unlockExcel';

export const runtime = 'edge';

export async function POST(req: NextRequest) {
    try {
        const contentType = req.headers.get('content-type') || '';
        if (!contentType.includes('multipart/form-data')) {
            return new Response(JSON.stringify({ error: 'Expected multipart/form-data' }), { status: 400 });
        }

        const formData = await req.formData();
        const files = formData.getAll('files[]');
        if (!files || files.length === 0) {
            return new Response(JSON.stringify({ error: 'No files selected' }), { status: 400 });
        }

        const results: Array<{ filename: string; ok: boolean; error?: string; dataUrl?: string }> = [];
        for (const file of files) {
            if (!(file instanceof File)) {
                continue;
            }
            const filename = file.name || 'file.xlsx';
            try {
                const arrayBuffer = await file.arrayBuffer();
                const inputBuffer = Buffer.from(arrayBuffer);
                if (!isLikelyExcelZip(inputBuffer)) {
                    results.push({ filename, ok: false, error: 'Not a valid Excel file' });
                    continue;
                }
                const outputBuffer = await unlockExcelBuffer(inputBuffer);
                const base64 = Buffer.from(outputBuffer).toString('base64');
                const dataUrl = `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${base64}`;
                results.push({ filename, ok: true, dataUrl });
            } catch (e) {
                results.push({ filename, ok: false, error: 'Error processing file' });
            }
        }

        return new Response(JSON.stringify({ results }), {
            status: 200,
            headers: { 'content-type': 'application/json' },
        });
    } catch (e) {
        return new Response(JSON.stringify({ error: 'Unexpected error' }), { status: 500 });
    }
}


