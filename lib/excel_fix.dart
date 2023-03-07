import 'dart:io';

import 'package:excel/excel.dart';
import 'package:file_selector/file_selector.dart';

class NewCellData {
  final String value;
  final CellStyle? style;
  final String? tag;
  NewCellData({required this.value, this.tag, this.style});
}

class ExcelFix {
  final XFile file;
  late Excel excel;
  Set<String> cleardExcels = {};
  Map<String, List<NewCellData>> appending = {};

  ExcelFix(this.file);

  void fix() async {
    var bytes = await file.readAsBytes();
    var excel = Excel.decodeBytes(bytes);
    this.excel = excel;

    try {
      fixExcel(excel);

      String fileName = file.name.replaceAll(".xlsx", "-已处理.xlsx");
      String filePath = file.path.replaceAll(".xlsx", "-已处理.xlsx");

      final newBytes = excel.save(fileName: fileName)!;
      if (Platform.isMacOS || Platform.isWindows) {
        File(filePath)
          ..createSync(recursive: true)
          ..writeAsBytesSync(newBytes);
      }
    } catch (exception) {
      print(exception.toString());
    }
  }

  void fixExcel(Excel excel) async {
    action1();
    action2();

    appending.forEach((key, list) {
      tryClearTable(key);
      final table = excel.tables[key];
      if (table != null) {
        var iAdd = 3;
        String? tag;
        for (var i = 0; i < list.length; i++) {
          if (list[i].tag != null && tag != list[i].tag) {
            table.updateCell(
                CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + iAdd),
                list[i].tag!,
                cellStyle: CellStyle());
            tag = list[i].tag;
            iAdd += 1;
          }
          table.updateCell(
              CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: i + iAdd),
              list[i].value,
              cellStyle: list[i].style);
        }
      }
    });
  }

  void action1() {
    final t1 = excel.tables["投料明细"];
    final s1 = t1?.rows.first;
    final l1 = s1?.firstWhere(
        (element) => element?.value.toString().contains("流程卡号") == true,
        orElse: () => null);
    final l2 = s1?.firstWhere(
        (element) => element?.value.toString().contains("投料工序") == true,
        orElse: () => null);
    if (t1 != null && l1 != null && l2 != null) {
      for (var row = 1; row < t1.maxRows; row++) {
        final id = t1.cell(CellIndex.indexByColumnRow(
            columnIndex: l1.colIndex, rowIndex: row));
        final process = t1.cell(CellIndex.indexByColumnRow(
            columnIndex: l2.colIndex, rowIndex: row));
        addIdProcess(
            NewCellData(
              value: id.value.toString(),
              tag: "------投料明细------",
              style: CellStyle(),
            ),
            process.value.toString());
      }
    }
  }

  void action2() {
    final t1 = excel.tables["生产明细表"];
    final s1 = t1?.rows.first;
    final l1 = s1?.firstWhere(
        (element) => element?.value.toString().contains("流程卡号") == true,
        orElse: () => null);
    final l2 = s1?.firstWhere(
        (element) => element?.value.toString().contains("下道工序") == true,
        orElse: () => null);
    if (t1 != null && l1 != null && l2 != null) {
      for (var row = 1; row < t1.maxRows; row++) {
        final id = t1.cell(CellIndex.indexByColumnRow(
            columnIndex: l1.colIndex, rowIndex: row));
        final process = t1.cell(CellIndex.indexByColumnRow(
            columnIndex: l2.colIndex, rowIndex: row));
        addIdProcess(
            NewCellData(
              value: id.value.toString(),
              tag: "------生产明细表------",
              style: CellStyle(),
            ),
            process.value.toString());
      }
    }
  }

  void tryClearTable(String name) {
    if (cleardExcels.contains(name)) return;
    final table = excel.tables[name];
    if (table == null) return;
    cleardExcels.add(name);
    for (var i = table.maxRows; i >= 3; i--) {
      table.removeRow(i);
    }
  }

  void addIdProcess(NewCellData id, String process) {
    appending[process] = appending[process] ?? <NewCellData>[];
    appending[process]!.add(id);
  }
}
