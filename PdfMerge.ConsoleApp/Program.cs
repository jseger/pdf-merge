using iText.Forms;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Action;
using iText.Kernel.Pdf.Canvas.Draw;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace PdfMerge.ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Let's start merging!");

            if (File.Exists("newPdf.PDF")) {
                File.Delete("newPdF.PDF");
            }

            MemoryStream baos = new MemoryStream();
            PdfDocument pdfDoc = new PdfDocument(new PdfWriter(baos));
            Document doc = new Document(pdfDoc);

            // Initialize a resultant document outlines in order to copy outlines from the source documents.
            // Note that outlines still could be copied even if in destination document outlines
            // are not initialized, by using PdfMerger with mergeOutlines vakue set as true
            pdfDoc.InitializeOutlines();

            // Copier contains the additional logic to copy acroform fields to a new page.
            // PdfPageFormCopier uses some caching logic which can potentially improve performance
            // in case of the reusing of the same instance.
            PdfPageFormCopier formCopier = new PdfPageFormCopier();


            // Copy all merging file's pages to the result pdf file
            Dictionary<String, PdfDocument> filesToMerge =
                System.IO.Directory.EnumerateFiles(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "*.PDF")
                .Where(s => !s.EndsWith("newPdf.PDF", StringComparison.InvariantCultureIgnoreCase) && !s.EndsWith("toc.PDF", StringComparison.InvariantCultureIgnoreCase))
                .ToDictionary(f => Path.GetFileName(f), f => new PdfDocument(new PdfReader(f)));

            Dictionary<int, String> toc = new Dictionary<int, String>();
            int page = 1;
            for (int k = 0; k < 25; k++) {
                foreach (KeyValuePair<String, PdfDocument> entry in filesToMerge) {
                    PdfDocument srcDoc = entry.Value;
                    int numberOfPages = srcDoc.GetNumberOfPages();

                    toc.Add(page, entry.Key);

                    for (int i = 1; i <= numberOfPages; i++, page++) {
                        Text text = new Text(String.Format("Page {0}", page));
                        srcDoc.CopyPagesTo(i, i, pdfDoc, formCopier);

                        // Put the destination at the very first page of each merged document
                        if (i == 1) {
                            text.SetDestination("p" + page);
                        }

                        doc.Add(new Paragraph(text).SetFixedPosition(page, 549, 810, 40)
                            .SetMargin(0)
                            .SetMultipliedLeading(1));
                    }
                }
            }

            PdfDocument tocDoc = new PdfDocument(new PdfReader("toc.pdf"));
            tocDoc.CopyPagesTo(1, 1, pdfDoc, formCopier);
            //tocDoc.Close();

            // Create a table of contents
            float tocYCoordinate = 750;
            float tocXCoordinate = doc.GetLeftMargin();
            float tocWidth = pdfDoc.GetDefaultPageSize().GetWidth() - doc.GetLeftMargin() - doc.GetRightMargin();
            int numberOfTocPages = 1;
            int count = 0;
            foreach (KeyValuePair<int, String> entry in toc) {
                count++;
                Paragraph p = new Paragraph();
                p.AddTabStops(new TabStop(500, TabAlignment.LEFT, new DashedLine()));
                p.Add(entry.Value);
                p.Add(new Tab());
                p.Add(entry.Key.ToString());
                p.SetAction(PdfAction.CreateGoTo("p" + entry.Key));
                doc.Add(p.SetFixedPosition(pdfDoc.GetNumberOfPages(), tocXCoordinate, tocYCoordinate, tocWidth)
                    .SetMargin(0)
                    .SetMultipliedLeading(1));

                tocYCoordinate -= 20;

                if(tocYCoordinate <= 30 && count < toc.Count) {
                    tocDoc.CopyPagesTo(1, 1, pdfDoc, formCopier);
                    tocYCoordinate = 750;
                    numberOfTocPages += 1;
                }
            }

            tocDoc.Close();

            foreach (PdfDocument srcDoc in filesToMerge.Values) {
                srcDoc.Close();
            }

            doc.Close();

            PdfDocument resultDoc = new PdfDocument(new PdfWriter("newPdf.pdf"));
            PdfDocument srcPdfDoc = new PdfDocument(new PdfReader(new MemoryStream(baos.ToArray()),
                new ReaderProperties()));
            srcPdfDoc.InitializeOutlines();

            // Create a copy order list and set the page with a table of contents as the first page
            Console.WriteLine("Number of pages: " + srcPdfDoc.GetNumberOfPages());
            int tocPageNumber = srcPdfDoc.GetNumberOfPages() - numberOfTocPages + 1;
            List<int> copyPagesOrderList = new List<int>();
            for (int i = tocPageNumber; i <= srcPdfDoc.GetNumberOfPages(); i++) {
                copyPagesOrderList.Add(i);
            }
            for (int i = 1; i < tocPageNumber; i++) {
                copyPagesOrderList.Add(i);
            }

            srcPdfDoc.CopyPagesTo(copyPagesOrderList, resultDoc, formCopier);

            srcPdfDoc.Close();
            resultDoc.Close();
        }
    }
}
