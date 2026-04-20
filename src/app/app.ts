import { Component, computed, signal } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { MatIconModule } from '@angular/material/icon';
import * as xlsx from 'xlsx';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type MergeRow = Record<string, any>;

type GroupedData = Record<string, {
    master: MergeRow;
    details: MergeRow[];
}>;

@Component({
  selector: 'app-root',
  imports: [RouterOutlet, FormsModule, MatIconModule],
  templateUrl: './app.html',
  styleUrl: './app.css'
})
export class App {
  docxFile = signal<File | null>(null);
  xlsxFile = signal<File | null>(null);
  excelRows = signal<MergeRow[]>([]);
  columns = computed(() => {
    const rows = this.excelRows();
    if (rows.length === 0) return [];
    return Object.keys(rows[0]);
  });

  groupingKey = signal<string>('');
  isProcessing = signal<boolean>(false);
  processedCount = signal<number>(0);
  totalCount = signal<number>(0);
  errorMsg = signal<string | null>(null);

  onDocxUpload(event: Event) {
    const input = event.target as HTMLInputElement;
    if (input.files && input.files[0]) {
      this.docxFile.set(input.files[0]);
      this.errorMsg.set(null);
    }
  }

  onXlsxUpload(event: Event) {
    const input = event.target as HTMLInputElement;
    if (input.files && input.files[0]) {
      const file = input.files[0];
      this.xlsxFile.set(file);
      this.errorMsg.set(null);

      const reader = new FileReader();
      reader.onload = (e: ProgressEvent<FileReader>) => {
        const bstr = e.target?.result as string;
        const wb = xlsx.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = xlsx.utils.sheet_to_json(ws);
        this.excelRows.set(data as MergeRow[]);
        if (this.columns().length > 0 && !this.groupingKey()) {
          this.groupingKey.set(this.columns()[0]);
        }
      };
      reader.readAsBinaryString(file);
    }
  }

  async runMerge() {
    const docx = this.docxFile();
    const rows = this.excelRows();
    const key = this.groupingKey();

    if (!docx || rows.length === 0 || !key) {
      this.errorMsg.set('Please provide all files and select a grouping key.');
      return;
    }

    this.isProcessing.set(true);
    this.errorMsg.set(null);
    this.processedCount.set(0);

    try {
      // 1. Group the data
      const grouped: GroupedData = {};
      rows.forEach((row) => {
        const keyValue = String(row[key]);
        if (!grouped[keyValue]) {
          grouped[keyValue] = {
            master: row,
            details: [],
          };
        }
        grouped[keyValue].details.push(row);
      });

      const groupKeys = Object.keys(grouped);
      this.totalCount.set(groupKeys.length);

      const templateArrayBuffer = await docx.arrayBuffer();
      const outputZip = new JSZip();

      // 2. Generate reports
      for (const groupKey of groupKeys) {
        const data = grouped[groupKey];
        const templateData = {
          ...data.master,
          items: data.details,
          now: new Date().toLocaleDateString(),
          totalItems: data.details.length
        };

        const pizZip = new PizZip(templateArrayBuffer);
        const doc = new Docxtemplater(pizZip, {
          paragraphLoop: true,
          linebreaks: true,
        });

        doc.render(templateData);

        const out = doc.getZip().generate({
          type: 'blob',
          mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });

        outputZip.file(`${groupKey}_merged.docx`, out);
        this.processedCount.update((c) => c + 1);
      }

      // 3. Save as ZIP
      const content = await outputZip.generateAsync({ type: 'blob' });
      saveAs(content, 'master_detail_merge_results.zip');
    } catch (err: unknown) {
      console.error(err);
      const message = err instanceof Error ? err.message : 'Unknown error';
      this.errorMsg.set(`Error during merge: ${message}`);
    } finally {
      this.isProcessing.set(false);
    }
  }
}
