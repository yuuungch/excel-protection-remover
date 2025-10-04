import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import xpath from 'xpath';

const EXCEL_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

function ensureXmlDeclaration(xml: string): string {
    const hasDecl = xml.trimStart().startsWith('<?xml');
    return hasDecl ? xml : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${xml}`;
}

function removeNodes(xml: string, xpathExpr: string): string {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xml, 'text/xml');
    const select = xpath.useNamespaces({ main: EXCEL_NS });
    const nodes: Node[] = select(xpathExpr, doc) as unknown as Node[];
    for (const node of nodes) {
        const parent = node.parentNode as Node | null;
        if (parent) {
            parent.removeChild(node);
        }
    }
    const serializer = new XMLSerializer();
    const out = serializer.serializeToString(doc);
    return ensureXmlDeclaration(out);
}

export function isLikelyExcelZip(buffer: Buffer): boolean {
    // Check for ZIP local file header signature: PK\x03\x04
    return buffer.length >= 4 && buffer[0] === 0x50 && buffer[1] === 0x4b && buffer[2] === 0x03 && buffer[3] === 0x04;
}

export async function validateExcelZip(zip: JSZip): Promise<boolean> {
    // Required entries for xlsx
    const hasContentTypes = !!zip.file('[Content_Types].xml');
    const hasWorkbook = !!zip.file('xl/workbook.xml');
    return hasContentTypes && hasWorkbook;
}

export async function unlockExcelBuffer(inputBuffer: Buffer): Promise<Buffer> {
    if (!isLikelyExcelZip(inputBuffer)) {
        throw new Error('Invalid ZIP header');
    }
    const zip = await JSZip.loadAsync(inputBuffer);
    if (!(await validateExcelZip(zip))) {
        throw new Error('Missing required XLSX parts');
    }

    // Process worksheets: remove <sheetProtection/>
    const worksheetFiles = zip.file(/^xl\/worksheets\/.*\.xml$/);
    for (const f of worksheetFiles) {
        const xml = await f.async('string');
        const updated = removeNodes(xml, '//main:sheetProtection');
        zip.file(f.name, updated);
    }

    // Process workbook.xml: remove <workbookProtection/> and <fileSharing/>
    const workbookPath = 'xl/workbook.xml';
    const workbookXml = await zip.file(workbookPath)!.async('string');
    let updatedWorkbook = removeNodes(workbookXml, '//main:workbookProtection');
    updatedWorkbook = removeNodes(updatedWorkbook, '//main:fileSharing');
    zip.file(workbookPath, updatedWorkbook);

    // Repack
    const outBuffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
    return outBuffer as Buffer;
}


