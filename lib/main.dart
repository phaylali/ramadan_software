import 'dart:io';
import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart';
import 'package:google_fonts/google_fonts.dart';
import 'package:pdf/pdf.dart';
import 'package:pdf/widgets.dart' as pw;
import 'package:path_provider/path_provider.dart';
import 'package:printing/printing.dart';

void main() {
  runApp(const ExcelToPdfApp());
}

class ExcelToPdfApp extends StatelessWidget {
  const ExcelToPdfApp({Key? key}) : super(key: key);

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Excel to PDF Converter',
      theme: ThemeData(primarySwatch: Colors.blue),
      home: ExcelToPdfPage(),
    );
  }
}

class ExcelToPdfPage extends StatefulWidget {
  @override
  _ExcelToPdfPageState createState() => _ExcelToPdfPageState();
}

class _ExcelToPdfPageState extends State<ExcelToPdfPage> {
  File? _excelFile;
  List<String> _headers = [];
  List<List<dynamic>> _dataRows = [];

  Future<void> _pickExcelFile() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx', 'xls'],
    );

    if (result != null) {
      setState(() {
        _excelFile = File(result.files.single.path!);
        _readExcelFile();
      });
    }
  }

  void _readExcelFile() {
    var bytes = _excelFile!.readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);

    // Assuming first sheet
    var table = excel.tables.keys.first;
    var sheet = excel.tables[table];

    // First row as headers
    setState(() {
      _headers =
          sheet!.rows.first
              .map((cell) => cell?.value?.toString() ?? '')
              .toList();

      // Data rows (skipping the first row)
      _dataRows =
          sheet.rows
              .skip(1)
              .map(
                (row) =>
                    row.map((cell) => cell?.value?.toString() ?? '').toList(),
              )
              .toList();
    });
  }

  pw.Widget _buildCustomCell(String header, String value, int index) {
    // Custom design for each cell based on index
    switch (index) {
      case 0: // First column
        return pw.Text(
          '$header : $value',
          style: pw.TextStyle(
            fontSize: 14,
            color: PdfColors.blue,

            fontWeight: pw.FontWeight.bold,
          ),
          textDirection: pw.TextDirection.rtl,
        );
      case 1: // Second column
        return pw.Container(
          padding: pw.EdgeInsets.all(5),
          decoration: pw.BoxDecoration(
            border: pw.Border.all(color: PdfColors.grey),
          ),
          child: pw.Text(
            '$header : $value',
            style: pw.TextStyle(fontSize: 12, color: PdfColors.green),
            textDirection: pw.TextDirection.rtl,
          ),
        );
      case 2: // Third column
        return pw.Align(
          alignment: pw.Alignment.centerRight,
          child: pw.Text(
            '$header : $value',
            style: pw.TextStyle(fontSize: 10, fontStyle: pw.FontStyle.italic),
            textDirection: pw.TextDirection.rtl,
          ),
        );
      // Add more custom designs for additional columns
      default:
        return pw.Text('$header : $value', textDirection: pw.TextDirection.rtl);
    }
  }

  Future<void> _convertToPdf() async {
    if (_dataRows.isEmpty) {
      ScaffoldMessenger.of(
        context,
      ).showSnackBar(SnackBar(content: Text('No Excel data to convert')));
      return;
    }

    // Load Tajawal font
    final tajawalFont = await PdfGoogleFonts.tajawalRegular();
    final tajawalBold = await PdfGoogleFonts.tajawalBold();
    final pdf = pw.Document();

    // Ensure we create full pages even if less than 3 rows
    int pageCount = (_dataRows.length / 3).ceil();

    for (int page = 0; page < pageCount; page++) {
      pdf.addPage(
        pw.Page(
          pageFormat: PdfPageFormat.a4,
          margin: pw.EdgeInsets.zero,
          build: (pw.Context context) {
            return pw.Column(
              children: [
                // First row
                _buildCustomRowSection(
                  rowData:
                      _dataRows.length > page * 3 ? _dataRows[page * 3] : [],
                  headers: _headers,
                  font: tajawalBold,
                  fontSmall: tajawalFont,
                ),

                // Second row
                _buildCustomRowSection(
                  rowData:
                      _dataRows.length > page * 3 + 1
                          ? _dataRows[page * 3 + 1]
                          : [],
                  headers: _headers,
                  font: tajawalBold,
                  fontSmall: tajawalFont,
                ),

                // Third row
                _buildCustomRowSection(
                  rowData:
                      _dataRows.length > page * 3 + 2
                          ? _dataRows[page * 3 + 2]
                          : [],
                  headers: _headers,
                  font: tajawalBold,
                  fontSmall: tajawalFont,
                ),
              ],
            );
          },
        ),
      );
    }

    // Save PDF
    final output = await getTemporaryDirectory();
    final file = File('${output.path}/excel_converted.pdf');
    await file.writeAsBytes(await pdf.save());

    // Preview PDF
    Navigator.push(
      context,
      MaterialPageRoute(builder: (context) => PDFPreviewScreen(pdfFile: file)),
    );
  }

  pw.Widget _buildCustomRowSection({
    required List<dynamic> rowData,
    required List<String> headers,
    required pw.Font font,
    required pw.Font fontSmall,
  }) {
    return pw.Container(
      height: PdfPageFormat.a4.height / 3,
      width: double.infinity,
      decoration: pw.BoxDecoration(
        border: pw.Border(
          bottom: pw.BorderSide(style: pw.BorderStyle.dashed, width: 1),
        ),
      ),
      child:
          rowData.isEmpty
              ? pw.Container()
              : pw.Column(
                mainAxisAlignment: pw.MainAxisAlignment.spaceEvenly,

                children: [
                  pw.Row(
                    mainAxisAlignment: pw.MainAxisAlignment.spaceBetween,
                    children: [
                      pw.SizedBox(width: 20),

                      pw.Expanded(
                        child: pw.Text(
                          '${headers[0]} : ${rowData[0]}',
                          style: pw.TextStyle(font: font, fontSize: 20),
                          textDirection: pw.TextDirection.rtl,
                          textAlign: pw.TextAlign.start,
                        ),
                      ),
                      pw.Spacer(),
                      pw.Expanded(
                        child: pw.Text(
                          '${headers[1]} : ${rowData[1]}',
                          style: pw.TextStyle(font: font, fontSize: 20),
                          textDirection: pw.TextDirection.rtl,
                          textAlign: pw.TextAlign.end,
                        ),
                      ),
                      pw.SizedBox(width: 20),
                    ],
                  ),

                  // Second row: third cell
                  pw.Row(
                    mainAxisAlignment: pw.MainAxisAlignment.center,
                    children: [
                      pw.Text(
                        '${rowData[2]}',

                        style: pw.TextStyle(
                          font: font,
                          fontSize: 26,
                          fontWeight: pw.FontWeight.bold,
                        ),
                        textDirection: pw.TextDirection.rtl,
                      ),
                      pw.Text(
                        '${headers[2]} : ',

                        style: pw.TextStyle(
                          font: fontSmall,
                          fontSize: 26,
                          fontWeight: pw.FontWeight.bold,
                        ),
                        textDirection: pw.TextDirection.rtl,
                      ),
                    ],
                  ),

                  // Third row: fourth and fifth cells with space between
                  pw.Row(
                    mainAxisAlignment: pw.MainAxisAlignment.spaceBetween,
                    children: [
                      pw.SizedBox(width: 20),

                      pw.Text(
                        '${rowData[3]}',
                        style: pw.TextStyle(font: font, fontSize: 16),
                        textDirection: pw.TextDirection.rtl,
                        textAlign: pw.TextAlign.end,
                      ),
                      pw.Text(
                        '${headers[3]} : ',
                        style: pw.TextStyle(font: fontSmall, fontSize: 16),
                        textDirection: pw.TextDirection.rtl,
                        textAlign: pw.TextAlign.start,
                      ),
                      pw.Spacer(),

                      pw.Text(
                        '${rowData[4]}',
                        style: pw.TextStyle(font: font, fontSize: 16),
                        textDirection: pw.TextDirection.rtl,
                        textAlign: pw.TextAlign.end,
                      ),
                      pw.Text(
                        '${headers[4]} : ',
                        style: pw.TextStyle(font: fontSmall, fontSize: 16),
                        textDirection: pw.TextDirection.rtl,
                        textAlign: pw.TextAlign.start,
                      ),

                      pw.SizedBox(width: 20),
                    ],
                  ),

                  // Fourth row: sixth and seventh cells
                  pw.Row(
                    mainAxisAlignment: pw.MainAxisAlignment.center,
                    children: [
                      pw.SizedBox(width: 20),
                      pw.Text(
                        '${rowData[5]}',
                        style: pw.TextStyle(font: font, fontSize: 16),
                        textDirection: pw.TextDirection.rtl,
                        textAlign: pw.TextAlign.end,
                      ),
                      pw.Text(
                        '${headers[5]} : ',
                        style: pw.TextStyle(font: fontSmall, fontSize: 16),
                        textDirection: pw.TextDirection.rtl,
                        textAlign: pw.TextAlign.start,
                      ),
                      pw.Spacer(),
                      pw.Text(
                        '${rowData[6]}',
                        style: pw.TextStyle(font: font, fontSize: 16),
                        textDirection: pw.TextDirection.rtl,
                        textAlign: pw.TextAlign.end,
                      ),
                      pw.Text(
                        '${headers[6]} : ',
                        style: pw.TextStyle(font: fontSmall, fontSize: 16),
                        textDirection: pw.TextDirection.rtl,
                        textAlign: pw.TextAlign.start,
                      ),

                      pw.SizedBox(width: 20),
                    ],
                  ),
                ],
              ),
    );
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: Text('Excel to PDF Converter')),
      body: Center(
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          children: [
            ElevatedButton(
              onPressed: _pickExcelFile,
              child: Text('Pick Excel File'),
            ),
            SizedBox(height: 20),
            Text(
              _excelFile != null
                  ? _excelFile!.path.split('/').last
                  : 'No file selected',
              style: TextStyle(fontSize: 16),
            ),
            SizedBox(height: 20),
            ElevatedButton(
              onPressed: _convertToPdf,
              child: Text('Convert to PDF'),
            ),
          ],
        ),
      ),
    );
  }
}

class PDFPreviewScreen extends StatelessWidget {
  final File pdfFile;

  const PDFPreviewScreen({Key? key, required this.pdfFile}) : super(key: key);

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text('PDF Preview'),
        actions: [
          IconButton(
            icon: Icon(Icons.share),
            onPressed: () async {
              await Printing.sharePdf(bytes: await pdfFile.readAsBytes());
            },
          ),
        ],
      ),
      body: PdfPreview(
        build: (format) => pdfFile.readAsBytesSync(),
        allowPrinting: true,
        allowSharing: true,
      ),
    );
  }
}
