package com.example.xlsopener

import android.app.DatePickerDialog
import android.os.Bundle
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
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
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

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        val groups = listOf("24-МПО", "24-БУХ", "22-ПРО-1", "22-ПРО-2", "22-ПРО-3", "22-ПРО-4")
        spnNameGroup = findViewById(R.id.groupName)
        val arrayAdapterNameGroup = ArrayAdapter(this, android.R.layout.simple_list_item_1, groups)
        spnNameGroup.adapter = arrayAdapterNameGroup

        recyclerViewClass = findViewById(R.id.classRecyclerView)
        recyclerViewClass.layoutManager = LinearLayoutManager(this)

        recyclerViewReplacement = findViewById(R.id.replacementRecyclerView)
        recyclerViewReplacement.layoutManager = LinearLayoutManager(this)

        CoroutineScope(Dispatchers.IO).launch {
            parseWebsite(url)
        }

        datePicker = findViewById(R.id.datePickerTextView)

        datePicker.setOnClickListener() {
            DatePickerDialog(this, { _, year, month, day ->
                val calendar = Calendar.getInstance()
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
                datePicker.text = "$dayName"
                updateSchedule()
            }, 2025, 0, 1).show()

        }
    }

    fun updateSchedule() {
        if (!isFileDownloaded) {
            println("File not downloaded yet")
            return
        }

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
            var status = ""
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
                var fileName = ""
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
                for (row in sheet) {
                    val rowData = mutableListOf<String>()
                    for (cell in row) {
                        when (cell.cellType) {
                            CellType.STRING -> rowData.add(cell.stringCellValue)
                            CellType.NUMERIC -> rowData.add(cell.numericCellValue.toString())
                            else -> rowData.add("")
                        }
                    }
                    data.add(rowData)
                }
            }
        } catch (e: Exception) {
            e.printStackTrace()
        }
        return data
    }

    fun filterSchedule(data: List<List<String>>, group: String, day: String): List<String> {
        val filteredData = mutableListOf<String>()
        try {
            if (data.size > 11) {
                val groupIndex = data[10].indexOf(group)
                var dayIndex = 0
                for (i in 0..61) {
                    if (data[i].isNotEmpty() && data[i][0] == "$day") {
                        dayIndex = i
                        break
                    }
                }

                filteredData.add("1 пара:" + data[dayIndex][groupIndex])
                filteredData.add("2 пара:" + data[dayIndex + 2][groupIndex])
                filteredData.add("3 пара:" + data[dayIndex + 4][groupIndex])
                if (data[dayIndex + 6][groupIndex] != data[dayIndex + 7][groupIndex]) {
                    filteredData.add(data[dayIndex + 6][groupIndex])
                    filteredData.add(data[dayIndex + 7][groupIndex])
                } else filteredData.add("4 пара:" + data[dayIndex + 6][groupIndex])

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