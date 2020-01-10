var excel = new ActiveXObject("Excel.Application");
excel.Visible = true;

var workbooks = [
  '',
]

for (var i = 0; i < workbooks.length; i++) {
  excel.Workbooks.Open(workbooks[i]);
}
