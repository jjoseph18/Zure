<html>
    <head>
        <title>Invoice Table </title>
        <style type = "text/css">
            table {
                border-collapse: collapse;
                width: 100%;
                color: #eb4034;
                font-family: monospace;
                font-size: 25px;
                text-align: left;
            }
            th {
                background-color: #eb4034;
                color: white;
            }
            tr:nth-child(even) {background-color: #ededed}
        </style>

    </head>
    <body>
        <table>
            <tr>
                <th>Bill To</th>
                <th>Description</th>
                <th> Date </th>
                <th> Hours </th>
                <th> Amount </th>

            </tr>
            <?php
            $ServerName = "exceldbserver.database.windows.net";
            $ConnectionOptions = array("Database" => "exceldb","Uid" => "aderant", "PWD" =>"Ad3rant!123");
            $conn = sqlsrv_connect($ServerName, $ConnectionOptions);
            $sql = "SELECT * FROM Invoicing";
            $result = sqlsrv_query($conn, $sql);
            $row_count = sqlsrv_num_rows($result);

            if ($row_count >= 0) {
                while ($row = sqlsrv_fetch_array($result)) {
                    echo "<tr><td>" . $row["BillTo"] . "</td><td>" . $row["Descriptions"] . "</td><td>" . $row["Dates"] . "</td><td>" . $row["HoursWorked"] . "</td><td>" . $row["Amount"] . "</td><td>";
                }
            }
            else {
                echo "No Results";
            }
            sqlsrv_close($conn);
            ?>
        </table>
    </body>
</html>
