package com.example.monthlytodo;

import android.app.Activity;
import android.content.SharedPreferences;
import android.graphics.Color;
import android.graphics.Paint;
import android.os.Bundle;
import android.view.Gravity;
import android.view.View;
import android.widget.ArrayAdapter;
import android.widget.Button;
import android.widget.CheckBox;
import android.widget.EditText;
import android.widget.LinearLayout;
import android.widget.Spinner;
import android.widget.TextView;

import org.json.JSONArray;
import org.json.JSONObject;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.Locale;
import java.util.UUID;

public class MainActivity extends Activity {
    private static final String STORE = "monthly_todo";
    private static final String TASKS = "tasks";
    private final ArrayList<Task> tasks = new ArrayList<>();
    private LinearLayout taskList;
    private Spinner monthSpinner;
    private TextView summary;
    private EditText taskInput;
    private String selectedMonth;

    private static class Task {
        String id, title, month, frequency;
        boolean done;
        Task(String title, String month, String frequency) {
            this.id = UUID.randomUUID().toString();
            this.title = title;
            this.month = month;
            this.frequency = frequency;
        }
    }

    @Override
    public void onCreate(Bundle state) {
        super.onCreate(state);
        loadTasks();
        selectedMonth = currentMonth();
        buildScreen();
        renderTasks();
    }

    private void buildScreen() {
        LinearLayout root = new LinearLayout(this);
        root.setOrientation(LinearLayout.VERTICAL);
        root.setPadding(dp(20), dp(18), dp(20), dp(12));
        root.setBackgroundColor(Color.rgb(247, 243, 237));

        TextView eyebrow = text("MONTHLY ORGANIZER", 12, Color.rgb(217, 119, 69));
        eyebrow.setTypeface(null, 1);
        root.addView(eyebrow);
        TextView title = text("My To-Do List", 30, Color.rgb(23, 42, 58));
        title.setTypeface(null, 1);
        root.addView(title, new LinearLayout.LayoutParams(-1, dp(48)));

        LinearLayout addRow = new LinearLayout(this);
        addRow.setGravity(Gravity.CENTER_VERTICAL);
        taskInput = new EditText(this);
        taskInput.setHint("Add a task...");
        taskInput.setSingleLine(true);
        taskInput.setTextSize(16);
        addRow.addView(taskInput, new LinearLayout.LayoutParams(0, dp(54), 1));
        Button add = button("ADD");
        add.setOnClickListener(v -> addTask());
        addRow.addView(add, new LinearLayout.LayoutParams(dp(82), dp(50)));
        root.addView(addRow);

        LinearLayout filterRow = new LinearLayout(this);
        filterRow.setGravity(Gravity.CENTER_VERTICAL);
        TextView monthLabel = text("VIEW MONTH", 12, Color.DKGRAY);
        monthLabel.setTypeface(null, 1);
        filterRow.addView(monthLabel, new LinearLayout.LayoutParams(dp(100), dp(52)));
        monthSpinner = new Spinner(this);
        ArrayList<String> months = new ArrayList<>();
        Calendar date = Calendar.getInstance();
        for (int i = -2; i <= 10; i++) {
            Calendar item = (Calendar) date.clone();
            item.add(Calendar.MONTH, i);
            months.add(new SimpleDateFormat("yyyy-MM", Locale.US).format(item.getTime()));
        }
        ArrayAdapter<String> adapter = new ArrayAdapter<>(this, android.R.layout.simple_spinner_dropdown_item, months);
        monthSpinner.setAdapter(adapter);
        monthSpinner.setSelection(months.indexOf(selectedMonth));
        monthSpinner.setOnItemSelectedListener(new android.widget.AdapterView.OnItemSelectedListener() {
            public void onNothingSelected(android.widget.AdapterView<?> parent) { }
            public void onItemSelected(android.widget.AdapterView<?> parent, View view, int position, long id) {
                selectedMonth = (String) parent.getItemAtPosition(position);
                renderTasks();
            }
        });
        filterRow.addView(monthSpinner, new LinearLayout.LayoutParams(0, dp(52), 1));
        root.addView(filterRow);

        summary = text("", 14, Color.rgb(83, 96, 105));
        root.addView(summary, new LinearLayout.LayoutParams(-1, dp(34)));
        taskList = new LinearLayout(this);
        taskList.setOrientation(LinearLayout.VERTICAL);
        root.addView(taskList, new LinearLayout.LayoutParams(-1, 0, 1));
        setContentView(root);
    }

    private void renderTasks() {
        if (taskList == null) return;
        taskList.removeAllViews();
        ArrayList<Task> visible = new ArrayList<>();
        for (Task task : tasks) if (task.month.equals(selectedMonth)) visible.add(task);
        Collections.sort(visible, Comparator.comparing(task -> task.done));
        int completed = 0;
        for (Task task : visible) {
            if (task.done) completed++;
            LinearLayout row = new LinearLayout(this);
            row.setGravity(Gravity.CENTER_VERTICAL);
            row.setPadding(dp(10), dp(4), dp(8), dp(4));
            row.setBackgroundColor(task.done ? Color.rgb(225, 237, 226) : Color.WHITE);
            CheckBox check = new CheckBox(this);
            check.setChecked(task.done);
            check.setOnClickListener(v -> toggleTask(task));
            row.addView(check, new LinearLayout.LayoutParams(dp(48), dp(58)));
            TextView label = text(task.title + "\n" + task.frequency, 16, Color.rgb(23, 42, 58));
            label.setGravity(Gravity.CENTER_VERTICAL);
            if (task.done) label.setPaintFlags(label.getPaintFlags() | Paint.STRIKE_THRU_TEXT_FLAG);
            row.addView(label, new LinearLayout.LayoutParams(0, dp(66), 1));
            Button delete = button("DELETE");
            delete.setTextSize(10);
            delete.setOnClickListener(v -> { tasks.remove(task); saveTasks(); renderTasks(); });
            row.addView(delete, new LinearLayout.LayoutParams(dp(72), dp(42)));
            LinearLayout.LayoutParams rowParams = new LinearLayout.LayoutParams(-1, dp(74));
            rowParams.setMargins(0, 0, 0, dp(8));
            taskList.addView(row, rowParams);
        }
        summary.setText(visible.size() + " tasks  ·  " + completed + " completed  ·  " + selectedMonth);
        if (visible.isEmpty()) {
            TextView empty = text("Nothing planned for this month yet.", 16, Color.rgb(83, 96, 105));
            empty.setGravity(Gravity.CENTER);
            taskList.addView(empty, new LinearLayout.LayoutParams(-1, dp(120)));
        }
    }

    private void addTask() {
        String title = taskInput.getText().toString().trim();
        if (title.isEmpty()) return;
        tasks.add(new Task(title, selectedMonth, "Monthly"));
        taskInput.setText("");
        saveTasks();
        renderTasks();
    }

    private void toggleTask(Task task) {
        task.done = !task.done;
        if (task.done) {
            String next = addMonths(task.month, task.frequency.equals("Quarterly") ? 3 : 1);
            boolean exists = false;
            for (Task candidate : tasks) {
                if (candidate.title.equals(task.title) && candidate.month.equals(next)) exists = true;
            }
            if (!exists) tasks.add(new Task(task.title, next, task.frequency));
        }
        saveTasks();
        renderTasks();
    }

    private void loadTasks() {
        try {
            JSONArray data = new JSONArray(getSharedPreferences(STORE, 0).getString(TASKS, "[]"));
            for (int i = 0; i < data.length(); i++) {
                JSONObject value = data.getJSONObject(i);
                Task task = new Task(value.getString("title"), value.getString("month"), value.optString("frequency", "Monthly"));
                task.id = value.optString("id", task.id);
                task.done = value.optBoolean("done", false);
                tasks.add(task);
            }
        } catch (Exception ignored) { }
    }

    private void saveTasks() {
        try {
            JSONArray data = new JSONArray();
            for (Task task : tasks) {
                JSONObject value = new JSONObject();
                value.put("id", task.id);
                value.put("title", task.title);
                value.put("month", task.month);
                value.put("frequency", task.frequency);
                value.put("done", task.done);
                data.put(value);
            }
            getSharedPreferences(STORE, 0).edit().putString(TASKS, data.toString()).apply();
        } catch (Exception ignored) { }
    }

    private String currentMonth() { return new SimpleDateFormat("yyyy-MM", Locale.US).format(new Date()); }

    private String addMonths(String value, int count) {
        try {
            Calendar date = Calendar.getInstance();
            date.setTime(new SimpleDateFormat("yyyy-MM", Locale.US).parse(value));
            date.add(Calendar.MONTH, count);
            return new SimpleDateFormat("yyyy-MM", Locale.US).format(date.getTime());
        } catch (Exception error) { return value; }
    }

    private TextView text(String value, int size, int color) {
        TextView view = new TextView(this);
        view.setText(value);
        view.setTextSize(size);
        view.setTextColor(color);
        return view;
    }

    private Button button(String label) {
        Button view = new Button(this);
        view.setText(label);
        view.setTextColor(Color.WHITE);
        view.setBackgroundColor(Color.rgb(217, 119, 69));
        return view;
    }

    private int dp(int value) { return (int) (value * getResources().getDisplayMetrics().density + 0.5f); }
}
