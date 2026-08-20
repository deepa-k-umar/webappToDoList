# Monthly Todo Website

## Run

From this folder, run:

```text
python server.py
```

Then open `http://localhost:8000`.

The website stores its working data in the browser. Clicking **Export Excel** sends all browser tasks to `MyToDoList.xlsx` in this same folder. Existing rows are preserved and task IDs prevent duplicate rows when exporting more than once.