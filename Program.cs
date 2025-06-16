using System;
using System.IO; // Explicitly use System.IO for Path operations
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing; // For Text, Paragraph, Run (used below for shapes)
using System.Collections.Generic; // For List<Dictionary<string, string>>
using CsvHelper; // For CSV parsing
using System.Globalization; // For CultureInfo.InvariantCulture
using System.Diagnostics; // For Process operations

namespace PptxSimpleGenerator
{
    class Program
    {
        // --- Configuration LibreOffice ---
        // Path to the LibreOffice executable (soffice.exe)
        // Ensure LibreOffice is installed on your Windows machine.
        // You might need to adjust this path based on your LibreOffice installation.
        // Example: private const string LIBREOFFICE_PATH = @"C:\Program Files\LibreOffice\program\soffice.exe";
        // If 'soffice' is in your system's PATH, you can use just "soffice".
        private const string LIBREOFFICE_PATH = "soffice";


        static void Main(string[] args)
        {
            Console.WriteLine("Démarrage du script de génération de présentation PPTX et PDF.");

            // --- Configuration ---
            // Path to your template PPTX file
            // Make sure 'template.pptx' is in the same directory as the executable
            // of this program, or adjust the path.
            string templatePptxPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "template.pptx");
            //Category of certifcate
            const string CATEGORY_CERT = "Orga";


            // Path to your CSV data file
            // Make sure 'data.csv' is in the same directory as the executable
            // or adjust the path.
            string csvDataPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data.csv");

            // Folder where the generated PPTX files will be saved
            // This folder will be created if it does not exist.
            string outputPptxFolder = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"generated_{CATEGORY_CERT}_pptx");

            // Folder where the generated PDF files will be saved
            // This folder will be created if it does not exist.
            string outputPdfFolder = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"generated_{CATEGORY_CERT}_pdf");

            // --- Mapping of PPTX Tags to CSV Headers ---
            // This is where you manually specify which tag in your PPTX
            // corresponds to which column header in your CSV file.
            // The dictionary key is the EXACT tag in your PPTX (e.g., "[[CLIENT_NAME]]").
            // The value is the EXACT column header in your CSV file (e.g., "CLIENT_NAME").
            Dictionary<string, string> tagToCsvHeaderMapping = new Dictionary<string, string>
            {
                { "[[VOTRE_BALISE]]", "NOM_COMPLET" },
                { "[[SUJET]]", "AUTRE" },
                // Add other mappings here based on your tags and CSV headers
                // Example: { "[[OTHER_TAG]]", "OTHER_CSV_HEADER" },
            };

            // --- Initial Checks ---
            if (!System.IO.File.Exists(templatePptxPath))
            {
                Console.WriteLine($"Fatal Error : Template PPTX file does not exist at : {templatePptxPath}");
                Console.WriteLine("Please ensure 'template.pptx' is present in the same folder as the executable.");
                Console.ReadKey(); // Wait for a key press before exiting
                return;
            }

            if (!System.IO.File.Exists(csvDataPath))
            {
                Console.WriteLine($"Fatal Error : CSV file does not exist at : {csvDataPath}");
                Console.WriteLine("Please ensure 'data.csv' is present in the same folder as the executable.");
                Console.ReadKey();
                return;
            }

            // Create output folders if they do not exist
            if (!System.IO.Directory.Exists(outputPptxFolder))
            {
                try
                {
                    System.IO.Directory.CreateDirectory(outputPptxFolder);
                    Console.WriteLine($"Output folder created : {outputPptxFolder}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Fatal Error : Unable to create output folder : {outputPptxFolder}");
                    Console.WriteLine($"Error details : {ex.Message}");
                    Console.WriteLine("Please check write permissions.");
                    Console.ReadKey();
                    return;
                }
            }

            if (!System.IO.Directory.Exists(outputPdfFolder))
            {
                try
                {
                    System.IO.Directory.CreateDirectory(outputPdfFolder);
                    Console.WriteLine($"PDF output folder created : {outputPdfFolder}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Fatal Error : Unable to create PDF output folder : {outputPdfFolder}");
                    Console.WriteLine($"Error details : {ex.Message}");
                    Console.WriteLine("Please check write permissions.");
                    Console.ReadKey();
                    return;
                }
            }

            Console.WriteLine("Starting CSV file read...");

            // --- Read CSV File ---
            List<Dictionary<string, string>> csvData = new List<Dictionary<string, string>>();
            try
            {
                using (var reader = new StreamReader(csvDataPath))
                using (var csv = new CsvHelper.CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    csv.Read();
                    csv.ReadHeader(); // Reads the first line as headers

                    Console.WriteLine("CSV headers detected : " + string.Join(", ", csv.HeaderRecord.Select(h => h ?? string.Empty)));

                    while (csv.Read())
                    {
                        var rowData = new Dictionary<string, string>();
                        foreach (var header in csv.HeaderRecord)
                        {
                            if (header != null)
                            {
                                rowData[header] = csv.GetField(header) ?? string.Empty;
                            }
                        }
                        csvData.Add(rowData);
                    }
                }
                Console.WriteLine($"{csvData.Count} CSV data rows read.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal Error while reading CSV file : {ex.Message}");
                Console.WriteLine($"Details : {ex.ToString()}");
                Console.ReadKey();
                return;
            }

            if (csvData.Count == 0)
            {
                Console.WriteLine("No valid data found in the CSV file. No presentations will be generated.");
                Console.ReadKey();
                return;
            }

            Console.WriteLine("Starting presentation generation...");

            // --- Iterate over each CSV row and generate PPTX and PDF ---
            for (int i = 0; i < csvData.Count; i++)
            {
                var rowData = csvData[i];
                string baseFilename = $"presentation_generee_{i + 1}";
                string currentPptxOutputFilename = $"{baseFilename}.pptx";
                string currentPptxOutputPath = System.IO.Path.Combine(outputPptxFolder, currentPptxOutputFilename);
                string currentPdfOutputFilename = $"{baseFilename}.pdf";
                string currentPdfOutputPath = System.IO.Path.Combine(outputPdfFolder, currentPdfOutputFilename);
                string tempCopyPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString() + ".pptx");

                Console.WriteLine($"Processing row {i + 1}/{csvData.Count}...");

                try
                {
                    // Step 1: Copy the template file for this iteration
                    System.IO.File.Copy(templatePptxPath, tempCopyPath, true);

                    // Step 2: Open the temporary copy for modification
                    using (PresentationDocument presentationDocument = PresentationDocument.Open(tempCopyPath, true))
                    {
                        // Iterate through all slides
                        foreach (SlidePart slidePart in presentationDocument.PresentationPart.SlideParts)
                        {
                            // We are specifically looking for text shapes (Drawing.Text)
                            var textElements = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>();

                            foreach (DocumentFormat.OpenXml.Drawing.Text textElement in textElements)
                            {
                                string currentTextContent = textElement.Text; // Get the initial text of the element
                                string modifiedTextContent = currentTextContent; // Variable to accumulate modifications

                                // Iterate through the defined mapping to replace tags
                                foreach (var mappingEntry in tagToCsvHeaderMapping)
                                {
                                    string pptxTag = mappingEntry.Key;    // Ex: "[[CLIENT_NAME]]"
                                    string csvHeader = mappingEntry.Value; // Ex: "CLIENT_NAME"

                                    // If the CSV header exists in the current data row
                                    if (rowData.ContainsKey(csvHeader))
                                    {
                                        string replacementValue = rowData[csvHeader]; // Get the value from the CSV

                                        // If the tag is found in the shape's text, replace it
                                        if (modifiedTextContent.Contains(pptxTag))
                                        {
                                            modifiedTextContent = modifiedTextContent.Replace(pptxTag, replacementValue);
                                        }
                                    }
                                }
                                textElement.Text = modifiedTextContent; // Update the element's text with all modifications
                            }
                        }

                        // Save the modifications to the temporary file
                        presentationDocument.Save();
                    }

                    // Step 3: Move the modified PPTX file to the final output folder
                    System.IO.File.Move(tempCopyPath, currentPptxOutputPath, true); // Move and overwrite if exists
                    Console.WriteLine($"Generated Presentation : {currentPptxOutputPath}");

                    // Step 4: Convert the generated PPTX to PDF using LibreOffice
                    try
                    {
                        var startInfo = new ProcessStartInfo
                        {
                            FileName = LIBREOFFICE_PATH,
                            Arguments = $"--headless --convert-to pdf \"{currentPptxOutputPath}\" --outdir \"{outputPdfFolder}\"",
                            UseShellExecute = false,
                            RedirectStandardOutput = true,
                            RedirectStandardError = true,
                            CreateNoWindow = true // Do not open a new window for LibreOffice
                        };

                        using (var process = Process.Start(startInfo))
                        {
                            if (process != null)
                            {
                                process.WaitForExit(); // Wait for LibreOffice process to exit

                                string stdout = process.StandardOutput.ReadToEnd();
                                string stderr = process.StandardError.ReadToEnd();

                                if (process.ExitCode != 0)
                                {
                                    Console.WriteLine($"Error converting PDF for {currentPptxOutputFilename}:");
                                    Console.WriteLine($"Stdout: {stdout}");
                                    Console.WriteLine($"Stderr: {stderr}");
                                    throw new Exception($"PDF conversion failed for {currentPptxOutputFilename}. LibreOffice Exit Code: {process.ExitCode}");
                                }
                                Console.WriteLine($"Generated PDF : {currentPdfOutputPath}");
                            }
                        }

                        if (!System.IO.File.Exists(currentPdfOutputPath))
                        {
                            Console.WriteLine($"Warning: PDF file was expected at {currentPdfOutputPath} but was not found after conversion.");
                        }
                    }
                    catch (System.ComponentModel.Win32Exception ex) when (ex.NativeErrorCode == 2) // Error code 2 means file not found (LibreOffice executable)
                    {
                        Console.WriteLine($"Fatal Error : LibreOffice executable not found at '{LIBREOFFICE_PATH}'.");
                        Console.WriteLine("Please ensure LibreOffice is installed and its 'program' directory is in your system's PATH, or set LIBREOFFICE_PATH to its full path.");
                        Console.WriteLine($"Error details : {ex.Message}");
                        return; // Exit the application if LibreOffice is not found
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error during PDF conversion for {currentPptxOutputFilename}: {ex.Message}");
                        Console.WriteLine($"Details : {ex.ToString()}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during presentation generation for row {i + 1} : {ex.Message}");
                    Console.WriteLine($"Details : {ex.ToString()}");
                }
                finally
                {
                    // Ensure the temporary PPTX file is deleted even if an error occurs
                    if (System.IO.File.Exists(tempCopyPath))
                    {
                        System.IO.File.Delete(tempCopyPath);
                    }
                }
            }

            Console.WriteLine("Generation process complete. Press any key to exit.");
            Console.ReadKey(); // Wait for a key press before exiting
        }
    }
}
