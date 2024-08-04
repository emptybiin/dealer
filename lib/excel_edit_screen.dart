import 'package:flutter/material.dart';
import 'package:path_provider/path_provider.dart';
import 'package:excel/excel.dart';
import 'dart:io';
import 'package:share/share.dart';

class ExcelEditScreen extends StatefulWidget {
  @override
  _ExcelEditScreenState createState() => _ExcelEditScreenState();
}

class _ExcelEditScreenState extends State<ExcelEditScreen> {
  final List<TextEditingController> _controllers = List.generate(21, (_) => TextEditingController());
  File? _excelFile;

  @override
  void initState() {
    super.initState();
    _loadExcelFile();
  }

  Future<void> _loadExcelFile() async {
    final directory = await getApplicationDocumentsDirectory();
    final path = directory.path;
    _excelFile = File('$path/report_copy.xlsx');
  }

  Future<void> _updateExcelFile() async {
    if (_excelFile != null) {
      var bytes = _excelFile!.readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);

      var sheet = excel['Sheet1'];

      // 문자열 셀 입력
      sheet.cell(CellIndex.indexByString("A6")).value = TextCellValue(_controllers[0].text);  // 고객명
      sheet.cell(CellIndex.indexByString("C9")).value = TextCellValue(_controllers[1].text);  // 차량명
      sheet.cell(CellIndex.indexByString("C10")).value = TextCellValue(_controllers[2].text); // 리스사명

      // 정수 셀 입력 및 계산
      int c11Value = int.parse(_controllers[3].text);
      int c12Value = int.parse(_controllers[4].text);
      sheet.cell(CellIndex.indexByString("C11")).value = TextCellValue('${c11Value}개월');  // 실행횟수 '개월'
      sheet.cell(CellIndex.indexByString("C12")).value = TextCellValue('${c12Value}개월');  // 납입횟수 '개월'

      sheet.cell(CellIndex.indexByString("I10")).value = IntCellValue(int.parse(_controllers[5].text));  // 차량매매가
      sheet.cell(CellIndex.indexByString("A17")).value = TextCellValue(_controllers[6].text.replaceFirstMapped(
          RegExp(r'(\d{2})(\d{2})(\d{2})'), (m) => '${m[1]}.${m[2]}.${m[3]}'));  // 최종납입 날짜 (YYMMDD 형식)

      sheet.cell(CellIndex.indexByString("I16")).value = IntCellValue(int.parse(_controllers[7].text));  // 미회수원금
      sheet.cell(CellIndex.indexByString("I17")).value = IntCellValue(int.parse(_controllers[8].text));  // 보증금

      // I18 계산: 선납금 * (실행횟수 - 납입횟수)
      int i18Value = int.parse(_controllers[9].text) * (c11Value - c12Value);
      sheet.cell(CellIndex.indexByString("I18")).value = IntCellValue(i18Value);  // 잔여선납금

      // I19 셀 입력
      sheet.cell(CellIndex.indexByString("I19")).value = IntCellValue(int.parse(_controllers[10].text));  // 잔존가치

      // I21 계산: (I16 - I17 - I18)
      int i16Value = int.parse(_controllers[7].text);
      int i17Value = int.parse(_controllers[8].text);
      sheet.cell(CellIndex.indexByString("I21")).value = IntCellValue(i16Value - i17Value - i18Value);  // 실미회수원금

      sheet.cell(CellIndex.indexByString("I23")).value = IntCellValue(int.parse(_controllers[11].text));  // 리스료

      // I24 계산: (C11 - C12)
      sheet.cell(CellIndex.indexByString("I24")).value = TextCellValue('${c11Value - c12Value}개월');  // 잔여회차

      sheet.cell(CellIndex.indexByString("I27")).value = IntCellValue(int.parse(_controllers[12].text));  // 일할차세
      sheet.cell(CellIndex.indexByString("I28")).value = IntCellValue(int.parse(_controllers[13].text));  // 일할이자
      sheet.cell(CellIndex.indexByString("I29")).value = IntCellValue(int.parse(_controllers[14].text));  // 승계수수료
      sheet.cell(CellIndex.indexByString("I30")).value = IntCellValue(int.parse(_controllers[15].text));  // 판매수수료
      sheet.cell(CellIndex.indexByString("I31")).value = IntCellValue(int.parse(_controllers[16].text));  // 기타비용

      // 추가된 문자열 셀 입력
      sheet.cell(CellIndex.indexByString("A32")).value = TextCellValue(_controllers[17].text);  // 추가입력 1
      sheet.cell(CellIndex.indexByString("I32")).value = IntCellValue(int.parse(_controllers[18].text));  // 추가입력 1 내용
      sheet.cell(CellIndex.indexByString("A33")).value = TextCellValue(_controllers[19].text);  // 추가입력 2
      sheet.cell(CellIndex.indexByString("I33")).value = IntCellValue(int.parse(_controllers[20].text));  // 추가입력 2 내용

      // I35 계산: I10 - I21 - I27 - I28 + I29 + I30 + I31 + I32 + I33
      int i10Value = int.parse(_controllers[5].text);
      int i29Value = int.parse(_controllers[14].text);
      int i30Value = int.parse(_controllers[15].text);
      int i31Value = int.parse(_controllers[16].text);
      int i32Value = int.parse(_controllers[18].text);
      int i33Value = int.parse(_controllers[20].text);
      sheet.cell(CellIndex.indexByString("I35")).value = IntCellValue(
          i10Value - (i16Value - i17Value - i18Value) - int.parse(_controllers[12].text) - int.parse(_controllers[13].text) + i29Value + i30Value + i31Value + i32Value + i33Value);  // 최종정산금액

      var encoded = excel.encode();
      if (encoded != null) {
        File(_excelFile!.path)
          ..createSync(recursive: true)
          ..writeAsBytesSync(encoded);
      }
    }
  }

  Future<void> _shareExcelFile() async {
    if (_excelFile != null) {
      Share.shareFiles([_excelFile!.path], text: 'Updated Excel file');
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text('Edit Excel'),
      ),
      body: Padding(
        padding: EdgeInsets.all(16.0),
        child: ListView(
          children: [
            for (int i = 0; i < _controllers.length; i++)
              Padding(
                padding: const EdgeInsets.symmetric(vertical: 8.0),
                child: TextField(
                  controller: _controllers[i],
                  decoration: InputDecoration(
                    labelText: _getLabelText(i),
                  ),
                  keyboardType: _getKeyboardType(i),  // 입력 형식에 맞게 키보드 타입 설정
                ),
              ),
            ElevatedButton(
              onPressed: () async {
                await _updateExcelFile();
              },
              child: Text('확인'),
            ),
            SizedBox(height: 20),
            ElevatedButton(
              onPressed: () async {
                await _shareExcelFile();
              },
              child: Text('공유'),
            ),
          ],
        ),
      ),
    );
  }

  String _getLabelText(int index) {
    switch (index) {
      case 0:
        return '고객명';
      case 1:
        return '차량명';
      case 2:
        return '리스사명';
      case 3:
        return '실행횟수';
      case 4:
        return '납입횟수';
      case 5:
        return '차량매매가';
      case 6:
        return '최종납입 날짜 (YYMMDD 형식)';
      case 7:
        return '미회수원금';
      case 8:
        return '보증금';
      case 9:
        return '선납금';
      case 10:
        return '잔존가치';
      case 11:
        return '리스료';
      case 12:
        return '일할차세';
      case 13:
        return '일할이자';
      case 14:
        return '승계수수료';
      case 15:
        return '판매수수료';
      case 16:
        return '기타비용';
      case 17:
        return '추가 입력 1';
      case 18:
        return '추가 입력 1 내용';
      case 19:
        return '추가 입력 2';
      case 20:
        return '추가 입력 2 내용';
      default:
        return 'Enter value for field ${index + 1}';
    }
  }

  TextInputType _getKeyboardType(int index) {
    if ([3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 18, 20].contains(index)) {
      return TextInputType.number;
    }
    return TextInputType.text;
  }
}
