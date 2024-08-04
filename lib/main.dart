import 'package:flutter/material.dart';
import 'package:path_provider/path_provider.dart';
import 'package:flutter/services.dart' show rootBundle;
import 'dart:io';
import 'excel_edit_screen.dart';

void main() {
  runApp(MyApp());
}

class MyApp extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      title: 'Excel App',
      theme: ThemeData(
        primarySwatch: Colors.blue,
      ),
      home: HomeScreen(),
    );
  }
}

class HomeScreen extends StatelessWidget {
  Future<File> copyExcelFile() async {
    try {
      final directory = await getApplicationDocumentsDirectory();
      final path = directory.path;
      final byteData = await rootBundle.load('assets/report.xlsx');
      final file = File('$path/report_copy.xlsx');
      await file.writeAsBytes(byteData.buffer.asUint8List());
      return file;
    } catch (e) {
      print('Error copying file: $e');
      throw e;
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text('Home'),
      ),
      body: Center(
        child: ElevatedButton(
          onPressed: () async {
            try {
              await copyExcelFile();
              Navigator.push(
                context,
                MaterialPageRoute(builder: (context) => ExcelEditScreen()),
              );
            } catch (e) {
              print('Error: $e');
            }
          },
          child: Text('엑셀'),
        ),
      ),
    );
  }
}
