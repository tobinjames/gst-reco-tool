declare module '*?raw' {
  const s: string;
  export default s;
}
declare module '*?inline' {
  const s: string;
  export default s;
}
declare module 'pdfjs-dist/build/pdf.worker.mjs';
interface ImportMeta {
  glob(pattern: string, options?: Record<string, unknown>): Record<string, unknown>;
}
