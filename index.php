<?php
$connect = new PDO("mysql:host=localhost;dbname=test","root", "");

$query = "SELECT * FROM tbl_customer ORDER BY CustomerID";

$statement = $connect->prepare($query);

$statement->execute();

$result = $statement->fetchAll();

$total_rows = $statement->rowCount();

$download_filelink = '<ul class="list-unstyled">';

if(isset($_POST["export"]))
{
    require_once 'class/PHPExcel.php';
    $last_page = ceil($total_rows/$_POST["records_no"]);

    $start = 0;
    $file_number = 0;
    for($count = 0; $count < $last_page; $count++)
    {
        $file_number++;
        $object = new PHPExcel();
        $object->setActiveSheetIndex(0);

        $table_columns = array("Nos", "Customer Name", "Gender", "Address", "City", "Postal Code", "Country");
        $column = 0;
        foreach($table_columns as $field)
        {
            $object->getActiveSheet()->setCellValueByColumnAndRow($column, 1, $field);
            $column++;
        }

        $query = "
  SELECT * FROM tbl_customer ORDER BY CustomerID LIMIT ".$start.", ".$_POST["records_no"]."
  ";
        $statement = $connect->prepare($query);
        $statement->execute();
        $excel_result = $statement->fetchAll();
        $excel_row = 2;
        foreach($excel_result as $sub_row)
        {
            $object->getActiveSheet()->setCellValueByColumnAndRow(0, $excel_row, $excel_row-1);
            $object->getActiveSheet()->setCellValueByColumnAndRow(1, $excel_row, $sub_row["CustomerName"]);
            $object->getActiveSheet()->setCellValueByColumnAndRow(2, $excel_row, $sub_row["Gender"]);
            $object->getActiveSheet()->setCellValueByColumnAndRow(3, $excel_row, $sub_row["Address"]);
            $object->getActiveSheet()->setCellValueByColumnAndRow(4, $excel_row, $sub_row["City"]);
            $object->getActiveSheet()->setCellValueByColumnAndRow(5, $excel_row, $sub_row["PostalCode"]);
            $object->getActiveSheet()->setCellValueByColumnAndRow(6, $excel_row, $sub_row["Country"]);
            $excel_row++;
        }
        $start = $start + $_POST["records_no"];
        $object_writer = PHPExcel_IOFactory::createWriter($object, 'Excel5');
        $file_name = 'File-'.$file_number.'.xls';
        $object_writer->save($file_name);
        $download_filelink .= '<li><label><a href="download.php?filename='.$file_name.'" target="_blank">Download - '.$file_name.'</a></label></li>';
    }
    $download_filelink .= '</ul>';
}


?>
<!doctype html>
<html>
<head>
    <title>Export Mysql Data into Multiple Excel File using PHP</title>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/2.2.0/jquery.min.js"></script>
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css" />
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/js/bootstrap.min.js"></script>
</head>
<body>
<div class="container box">
    <h3 align="center">Export Mysql Data into Multiple Excel File using PHP</h3>
    <br />
    <br />
    <form method="post">
        <div class="row">
            <div class="col-md-3" align="right"><label>No. of Records in Each File</label></div>
            <div class="col-md-2">
                <select name="records_no" class="form-control">
                    <option value="5">5 per file</option>
                    <option value="10">10 per file</option>
                    <option value="15">15 per file</option>
                </select>
            </div>
            <div class="col-md-2">
                <input type="submit" name="export" class="btn btn-success" value="Export to Excel" />
            </div>
            <div class="col-md-5">
                <?php echo $download_filelink; ?>
            </div>
        </div>
    </form>
    <br />
    <div class="table-responsive">
        <table id="customer_data" class="table table-bordered table-striped">
            <thead>
            <tr>
                <th>Customer Name</th>
                <th>Gender</th>
                <th>Address</th>
                <th>City</th>
                <th>Postal Code</th>
                <th>Country</th>
            </tr>
            </thead>
            <tbody>
            <?php
            foreach($result as $row)
            {
                echo '
      <tr>
       <td>'.$row["CustomerName"].'</td>
       <td>'.$row["Gender"].'</td>
       <td>'.$row["Address"].'</td>
       <td>'.$row["City"].'</td>
       <td>'.$row["PostalCode"].'</td>
       <td>'.$row["Country"].'</td>
      </tr>
      ';
            }
            ?>
            </tbody>
        </table>
    </div>
</div>
<br />
<br />
</body>
</html>