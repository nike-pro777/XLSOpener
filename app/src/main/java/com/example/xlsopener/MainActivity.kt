package com.example.xlsopener

import android.app.DatePickerDialog
import android.content.Intent
import android.os.Bundle
import android.view.Menu
import android.view.MenuItem
import android.widget.ArrayAdapter
import android.widget.Spinner
import android.widget.TextView
import androidx.appcompat.app.AppCompatActivity
import androidx.recyclerview.widget.LinearLayoutManager
import androidx.recyclerview.widget.RecyclerView
import kotlinx.coroutines.CoroutineScope
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import java.io.File
import java.io.FileOutputStream
import org.jsoup.Jsoup
import okhttp3.OkHttpClient
import okhttp3.Request
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.util.Calendar

class MainActivity : AppCompatActivity() {

    private val url = "https://college.tu-bryansk.ru/?page_id=4043"
    private lateinit var spnNameGroup: Spinner
    private lateinit var recyclerViewClass: RecyclerView
    private lateinit var recyclerViewReplacement: RecyclerView
    private var isFileDownloaded = false
    private lateinit var datePicker: TextView
    private lateinit var dateDay: String
    private lateinit var weekStateText: String

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        val _class = listOf("301н", "308н", "310н", "204н", "309", "306", "210")
        spnNameGroup = findViewById(R.id.groupName)
        val arrayAdapterNameGroup = ArrayAdapter(this, android.R.layout.simple_list_item_1, _class)
        spnNameGroup.adapter = arrayAdapterNameGroup

        recyclerViewClass = findViewById(R.id.classRecyclerView)
        recyclerViewClass.layoutManager = LinearLayoutManager(this)

        recyclerViewReplacement = findViewById(R.id.replacementRecyclerView)
        recyclerViewReplacement.layoutManager = LinearLayoutManager(this)

        if (!isFileDownloaded) {
            CoroutineScope(Dispatchers.IO).launch {
                parseWebsite(url)
            }
        }

        initDatePicker()
    }

    override fun onCreateOptionsMenu(menu: Menu): Boolean {
        menuInflater.inflate(R.menu.menu_main, menu)
        return true
    }

    override fun onOptionsItemSelected(item: MenuItem): Boolean {
        return when (item.itemId) {
            R.id.action_info -> {
                val intent = Intent(this, InfoActivity::class.java)
                startActivity(intent)
                true
            }

            else -> super.onOptionsItemSelected(item)
        }
    }

    private fun initDatePicker() {
        datePicker = findViewById(R.id.datePickerTextView)
        val calendar = Calendar.getInstance()
        val currentYear = calendar.get(Calendar.YEAR)
        val currentMonth = calendar.get(Calendar.MONTH)
        val currentDay = calendar.get(Calendar.DAY_OF_MONTH)
        datePicker.setOnClickListener() {
            DatePickerDialog(this, { _, year, month, day ->
                dateDay = "$day.${month + 1}.${year}г."
                calendar.set(year, month, day)
                val dayOfWeek = calendar.get(Calendar.DAY_OF_WEEK)
                val dayName = when (dayOfWeek) {
                    Calendar.SUNDAY -> "Воскресенье"
                    Calendar.MONDAY -> "Понедельник"
                    Calendar.TUESDAY -> "Вторник"
                    Calendar.WEDNESDAY -> "Среда"
                    Calendar.THURSDAY -> "Четверг"
                    Calendar.FRIDAY -> "Пятница"
                    Calendar.SATURDAY -> "Суббота"
                    else -> "Неизвестно"
                }
                datePicker.text = dayName
                updateSchedule()
            }, currentYear, currentMonth, currentDay).show()

        }
    }

    fun updateSchedule() {
        if (!isFileDownloaded) {
            println("File not downloaded yet")
            return
        }
        val sheduleHeadingText = findViewById<TextView>(R.id.headingTextShedule)
        sheduleHeadingText.text = "Список занятий ($weekStateText недели)"
        val group = spnNameGroup.selectedItem as String
        val day = datePicker.text.toString()
        val sheduleFile = File(getExternalFilesDir(null), "sheduleFile.xlsx")
        val replacementFile = File(getExternalFilesDir(null), "replacementFile.xlsx")

        if (sheduleFile.exists()) {
            val data = parseExcelFile(sheduleFile, 0)
            val filteredData = filterSchedule(data, group, day)
            recyclerViewClass.adapter = OrderAdapter(filteredData)
        } else {
            println("File not found")
        }

        if (replacementFile.exists()) {
            val data = parseExcelFile(replacementFile, 1)
            val filteredData = filterReplacement(data, group, dateDay)
            recyclerViewReplacement.adapter = ReplacementAdapter(filteredData)
        } else {
            println("File not found")
        }
    }

    private fun parseWebsite(url: String) {
        try {
            val document = Jsoup.connect(url).get()
            val links = document.select("a[href]")
            val weekStates = document.select("strong")
            for (weekState in weekStates) {
                if (weekState.text().contains("чётной")) {
                    weekStateText = weekState.text()
                }
            }

            if (File(getExternalFilesDir(null), "sheduleFile.xlsx").exists()
                && File(getExternalFilesDir(null), "replacementFile.xlsx").exists()
            ) {
                isFileDownloaded = true
            }

            var status: String
            for (link in links) {
                val linkText = link.text()
                if (linkText.contains("Расписание учебных занятий")) {
                    val href = link.attr("href")
                    println("Found file: $href")
                    status = "shedule"
                    downloadFile(href, status)
                }
                if (linkText.contains("Замены")) {
                    val href = link.attr("href")
                    println("Found file: $href")
                    status = "replacement"
                    downloadFile(href, status)
                }
            }
        } catch (e: Exception) {
            e.printStackTrace()
        }
    }

    private fun downloadFile(fileUrl: String, statusFile: String) {
        val client = OkHttpClient()
        val request = Request.Builder().url(fileUrl).build()

        try {
            val response = client.newCall(request).execute()
            if (response.isSuccessful) {
                val fileName: String
                if (statusFile == "shedule") {
                    fileName = "sheduleFile.xlsx"
                } else fileName = "replacementFile.xlsx"
                val file = File(getExternalFilesDir(null), fileName)
                FileOutputStream(file).use { outputStream ->
                    outputStream.write(response.body?.bytes())
                }
                println("File downloaded: ${file.absolutePath}")
                isFileDownloaded = true
            }
        } catch (e: Exception) {
            e.printStackTrace()
        }
    }

    fun parseExcelFile(file: File, sheetNumber: Int): List<List<String>> {
        val data = mutableListOf<List<String>>()
        try {
            FileInputStream(file).use { fis ->
                val workbook: Workbook = XSSFWorkbook(fis)
                val sheet: Sheet = workbook.getSheetAt(sheetNumber)
                val mergedRegions = sheet.mergedRegions // Получаем список объединенных областей

                for (rowIndex in 0..sheet.lastRowNum) {
                    val row = sheet.getRow(rowIndex) ?: continue // Пропускаем пустые строки
                    val rowData = mutableListOf<String>()

                    for (cellIndex in 0 until row.lastCellNum) {
                        val cell =
                            row.getCell(cellIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
                        val cellValue = getCellValue(cell, sheet, mergedRegions)
                        rowData.add(cellValue)
                    }
                    data.add(rowData)
                }
            }
        } catch (e: Exception) {
            e.printStackTrace()
        }
        return data
    }

    // Функция для получения значения ячейки с учетом объединенных областей
    fun getCellValue(cell: Cell, sheet: Sheet, mergedRegions: List<CellRangeAddress>): String {
        // Проверяем, находится ли ячейка в объединенной области
        for (mergedRegion in mergedRegions) {
            if (mergedRegion.isInRange(cell.rowIndex, cell.columnIndex)) {
                // Если ячейка объединена, берем значение из верхней левой ячейки области
                val firstRow = sheet.getRow(mergedRegion.firstRow) ?: return ""
                val firstCell = firstRow.getCell(mergedRegion.firstColumn) ?: return ""

                // Проверяем, чтобы не было рекурсии (если ячейка ссылается сама на себя)
                if (firstCell.rowIndex == cell.rowIndex && firstCell.columnIndex == cell.columnIndex) {
                    return when (cell.cellType) {
                        CellType.STRING -> cell.stringCellValue
                        CellType.NUMERIC -> cell.numericCellValue.toString()
                        else -> ""
                    }
                }
                return getCellValue(firstCell, sheet, mergedRegions) // Рекурсивно получаем значение
            }
        }

        // Если ячейка не объединена, возвращаем её значение
        return when (cell.cellType) {
            CellType.STRING -> cell.stringCellValue
            CellType.NUMERIC -> cell.numericCellValue.toString()
            else -> ""
        }
    }

    fun filterSchedule(data: List<List<String>>, classRoom: String, day: String): List<String> {
        val filteredData = mutableListOf<String>()

        try {
            if (data.size > 11) {
                val groupIndex = data[10].indexOf(classRoom)
                var dayIndex = 0
                for (i in 0..61) {
                    if (data[i].isNotEmpty() && data[i][0] == day) {
                        dayIndex = i
                        break
                    }
                }

                fun filterClass(classNumber: Int, firstNumber: Int, secondNumber: Int) {
                    for (j in 1..data[0].size) {
                        if (classNumber == 4 && data[dayIndex + firstNumber][j].contains("4п")
                            && (data[dayIndex + secondNumber][j].contains(classRoom) || data[dayIndex + firstNumber][j].contains(
                                classRoom
                            ))
                        ) {
                            filteredData.add("${data[dayIndex + firstNumber][j] + " " + data[10][j]} \n ${data[dayIndex + secondNumber][j] + " " + data[10][j]}")
                            return
                        }
                        if (data[dayIndex + firstNumber][j].isEmpty() && data[dayIndex + secondNumber][j].isNotEmpty()
                            && data[dayIndex + secondNumber][j].contains(classRoom)
                        ) {
                            filteredData.add("${classNumber}пара:" + data[dayIndex + secondNumber][j] + " " + data[10][j])
                        }
                        if (data[dayIndex + firstNumber][j].isNotEmpty() && data[dayIndex + secondNumber][j].isEmpty()
                            && data[dayIndex + firstNumber][j].contains(classRoom)
                        ) {
                            filteredData.add("${classNumber}пара:" + data[dayIndex + firstNumber][j] + " " + data[10][j])
                        }
                        if (data[dayIndex + firstNumber][j].isNotEmpty() && data[dayIndex + secondNumber][j].isNotEmpty()
                            && data[dayIndex + secondNumber][j] == data[dayIndex + firstNumber][j]
                            && data[dayIndex + secondNumber][j].contains(classRoom)
                        ) {
                            filteredData.add("${classNumber}пара:" + data[dayIndex + firstNumber][j] + " " + data[10][j])
                        }
                        if (data[dayIndex + firstNumber][j].isNotEmpty() && data[dayIndex + secondNumber][j].isNotEmpty()
                            && data[dayIndex + secondNumber][j] != data[dayIndex + firstNumber][j] && weekStateText == "чётной"
                            && (data[dayIndex + secondNumber][j].contains(classRoom))
                        ) {
                            filteredData.add("${classNumber}пара:" + data[dayIndex + secondNumber][j] + " " + data[10][j])
                        }
                        if (data[dayIndex + firstNumber][j].isNotEmpty() && data[dayIndex + secondNumber][j].isNotEmpty()
                            && data[dayIndex + secondNumber][j] != data[dayIndex + firstNumber][j] && weekStateText != "чётной"
                            && data[dayIndex + firstNumber][j].contains(classRoom)
                        ) {
                            filteredData.add("${classNumber}пара:" + data[dayIndex + firstNumber][j] + " " + data[10][j])
                        }
                    }
                }

                filterClass(1, 0, 1)
                filterClass(2, 2, 3)
                filterClass(3, 4, 5)
                filterClass(4, 6, 7)

            }
        } catch (e: Exception) {
            e.printStackTrace()
        }
        return filteredData
    }

    fun filterReplacement(data: List<List<String>>, group: String, day: String): List<String> {
        val filteredData = mutableListOf<String>()

        try {
            var dayIndex = -1
            for (i in data.indices) {
                if (data[i].isNotEmpty() && data[i][0].contains(day)) {
                    dayIndex = i
                    break
                }
            }

            if (dayIndex != -1) {
                var groupFound = false

                for (i in dayIndex + 3 until data.size) {
                    if (data[i][0].isEmpty() && data[i][1].isEmpty() && data[i][2].isEmpty()) {
                        break
                    }

                    if (data[i].isNotEmpty() && data[i][0] == group) {
                        groupFound = true
                    }

                    if (groupFound && (data[i][0] == group || data[i][0].isEmpty())) {
                        val para = data[i][1]
                        val insteadOf = data[i][2]
                        val replacement = data[i][3]

                        val replacementInfo = "$para: $insteadOf <-> $replacement"
                        filteredData.add(replacementInfo)
                    }
                }
            }
        } catch (e: Exception) {
            e.printStackTrace()
        }
        return filteredData
    }
}