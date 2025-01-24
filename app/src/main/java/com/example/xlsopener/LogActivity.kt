package com.example.xlsopener

import android.os.Bundle
import android.widget.ArrayAdapter
import androidx.activity.enableEdgeToEdge
import androidx.appcompat.app.AppCompatActivity
import androidx.core.view.ViewCompat
import androidx.core.view.WindowInsetsCompat
import androidx.recyclerview.widget.LinearLayoutManager
import androidx.recyclerview.widget.RecyclerView

class LogActivity : AppCompatActivity() {
    companion object {
        val logs : MutableList<String> = mutableListOf()
    }
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_log)

        val recyclerViewLog = findViewById<RecyclerView>(R.id.logRecyclerView)
        recyclerViewLog.layoutManager = LinearLayoutManager(this)

        recyclerViewLog.adapter = ReplacementAdapter(logs)


    }
}