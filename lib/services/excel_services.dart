import 'dart:math';

import 'package:excel/excel.dart';
import 'package:flutter_file_dialog/flutter_file_dialog.dart';
import 'package:path_provider/path_provider.dart';
import 'dart:io';

import '../models/task.dart';

Future<void> exportTasksToExcel(List<Task> tasks) async {
  List<String> colors = ["primaryColor", "red", "yellow", "black"];
  List<String> isComplete = ["Pending", "Completed"];

  // Create a new Excel file
  var excel = Excel.createExcel();

  // Access the 'Sheet1'
  Sheet sheet = excel['Sheet1'];

  // Column headers
  List<String> headers = [
    "Id",
    "Title",
    "Note",
    "Date",
    "StartTime",
    "EndTime",
    "Remind",
    "Repeat",
    "Color",
    "isCompleted",
    "createdAt",
    "updatedAt"
  ];
  sheet.appendRow(headers.cast<CellValue?>());

  // If you have a list of tasks, convert each task to a list and add it to the rows
  for (Task task in tasks) {
    List<String?> row = [
      task.id?.toString(),
      task.title,
      task.note,
      task.date,
      task.startTime,
      task.endTime,
      task.remind?.toString(),
      task.repeat,
      colors[task.color ?? 0],
      isComplete[task.isCompleted ?? 0],
      task.createdAt?.toString(),
      task.updatedAt?.toString(),
    ];
    sheet.appendRow(row.cast<CellValue?>());
  }

  // Save the Excel file
  final directory = await getTemporaryDirectory();
  final path = directory.path;
  final file = File('$path/tasks.xlsx');
  await file.writeAsBytes(excel.save() ?? <int>[]);

  var rand = Random();
  int randomNumber = rand.nextInt(50);

  final params = SaveFileDialogParams(
    sourceFilePath: file.path,
    localOnly: true,
    fileName: "Tasks$randomNumber.xlsx",
  );

  await FlutterFileDialog.saveFile(params: params);
}
