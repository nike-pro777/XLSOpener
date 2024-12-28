package com.example.xlsopener

import android.app.DatePickerDialog
import android.os.Bundle
import android.view.View
import android.widget.AdapterView
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
import okio.ByteString.Companion.encode
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.nio.charset.Charset

class MainActivity : AppCompatActivity() {

    private val url = "https://college.tu-bryansk.ru/?page_id=4043"
    private lateinit var spnNameGroup: Spinner
    private lateinit var recyclerViewClass: RecyclerView
    private var isFileDownloaded = false

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        val groups = listOf("22-ПРО-1", "22-ПРО-2", "22-ПРО-3", "22-ПРО-4")
        spnNameGroup = findViewById(R.id.groupName)
        val arrayAdapterNameGroup = ArrayAdapter(this, android.R.layout.simple_list_item_1, groups)
        spnNameGroup.adapter = arrayAdapterNameGroup

//        val dates = listOf("27.12.24", "26.12.24")
//        spnDate = findViewById(R.id.datePicker)
//        val arrayAdapterDate = ArrayAdapter(this, android.R.layout.simple_list_item_1, dates)
//        spnDate.adapter = arrayAdapterDate

        recyclerViewClass = findViewById(R.id.classRecyclerView)
        recyclerViewClass.layoutManager = LinearLayoutManager(this)

        CoroutineScope(Dispatchers.IO).launch {
            parseWebsite(url)
        }

        spnNameGroup.onItemSelectedListener = object : AdapterView.OnItemSelectedListener {
            override fun onItemSelected(parent: AdapterView<*>?, view: View?, position: Int, id: Long) {
                updateSchedule()
            }

            override fun onNothingSelected(parent: AdapterView<*>?) {}
        }

        val datePicker = findViewById<TextView>(R.id.datePickerTextView)

        datePicker.setOnClickListener() {
            DatePickerDialog(this,{_,year,month,day ->
                datePicker.text = "${month+1}.$day"
            },2024,12,31).show()
        }
    }

    fun updateSchedule() {
        if (!isFileDownloaded) {
            println("File not downloaded yet")
            return
        }
        val group = spnNameGroup.selectedItem as String
        val day = "Вторник"
        val file = File(getExternalFilesDir(null), "file.xlsx")

        if (file.exists()) {
            println("file exists")
            val data = parseExcelFile(file)
            val filteredData = filterSchedule(data, group, day)
            recyclerViewClass.adapter = OrderAdapter(filteredData)
        } else {
            println("File not found")
        }
    }

    private fun parseWebsite(url: String) {
        try {
            val document = Jsoup.connect(url).get()
            val links = document.select("a[href]")

            for (link in links) {
                val linkText = link.text()
                if (linkText.contains("Расписание учебных занятий")) {
                    val href = link.attr("href")
                    println("Found file: $href")
                    downloadFile(href)
                    break
                }
            }
        } catch (e: Exception) {
            e.printStackTrace()
        }
    }

    private fun downloadFile(fileUrl: String) {
        val client = OkHttpClient()
        val request = Request.Builder().url(fileUrl).build()

        try {
            val response = client.newCall(request).execute()
            if (response.isSuccessful) {
                val fileName = "file.xlsx"
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

    fun parseExcelFile(file: File): List<List<String>> {
        val data = mutableListOf<List<String>>()
        try {
            FileInputStream(file).use { fis ->
                val workbook: Workbook = XSSFWorkbook(fis)
                for(i in 0 until 1) {
                    val sheet: Sheet = workbook.getSheetAt(i)

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
                val groupIndex = data[11].indexOf(group)
                val dayIndex = data.indexOfFirst { it[1] == day }

                if (groupIndex != -1 && dayIndex != -1) {
                    for (i in dayIndex until data.size) {
                        if (data[i][1] != day) break
                        filteredData.add(data[i][groupIndex])
                    }
                }
            }
        } catch (e: Exception) {
            e.printStackTrace()
        }
        return filteredData
    }
}