<%@ page contentType="text/html;charset=UTF-8" pageEncoding="utf-8" trimDirectiveWhitespaces="true" session="false" %>
<%@ taglib prefix="c" uri="http://java.sun.com/jsp/jstl/core" %>
<%@ taglib prefix="form" uri="http://www.springframework.org/tags/form" %>

<!DOCTYPE html>
<html>
<head>
    <meta name="viewport" content="width=device-width" />
    <meta charset="utf-8" />
    <title>Create an Excel Workbook</title>
    <link href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">
    <style type="text/css">
        thead tr {
            text-align: center;
        }

        .table th {
            color: white;
            background-color: rgb(242, 101, 34);
            padding: 0.5rem;
        }

        .table td {
            padding: 0.5rem;
        }

        .table input {
            font-size: 0.9rem;
            padding: 0.3rem 0.35rem;
        }

        .first-column {
            text-align: center;
        }
    </style>
</head>
<body style="padding:20px; font-size: 0.9rem">

    <form:form action="${create}" method="POST" modelAttribute="workbookItemsWithFormat">
        <table class="table table-bordered table-condensed">
            <colgroup>
                <col style="width: 10%" />
                <col style="width: 45%" />
                <col style="width: 45%" />
            </colgroup>
            <thead>
                <tr>
                    <th>ID</th>
                    <th>First Name</th>
                    <th>Last Name</th>
                </tr>
            </thead>
            <tbody>

                <c:forEach items="${workbookItemsWithFormat.items}" varStatus="i">
                    <tr>
                        <td><form:input path="items[${i.index}].id" type="text" class="form-control-plaintext first-column" readonly="true" /></td>
                        <td><form:input path="items[${i.index}].firstName" type="text" class="form-control" /></td>
                        <td><form:input path="items[${i.index}].lastName" type="text" class="form-control"/></td>
                    </tr>
                </c:forEach>
            </tbody>
        </table>
        <div>
            <h4>Output format:</h4>
            <div class="form-check">
                <form:radiobutton path="selectedFormat" value="XLSX" class="form-check-input"/>
                <label for="XLSX" class="form-check-label">XLSX</label>
            </div>
            <div class="form-check">
                   <form:radiobutton path="selectedFormat" value="XLS" class="form-check-input"/>
                <label for="XLS" class="form-check-label">XLS</label>
            </div>
            <div class="form-check">
                <form:radiobutton path="selectedFormat" value="ODS" class="form-check-input"/>
                <label for="ODS" class="form-check-label">ODS</label>
            </div>
            <div class="form-check">
                <form:radiobutton path="selectedFormat" value="CSV" class="form-check-input"/>
                <label for="CSV" class="form-check-label">CSV</label>
            </div>
            <div class="form-check">
                <form:radiobutton path="selectedFormat" value="HTML" class="form-check-input"/>
                <label for="HTML" class="form-check-label">HTML</label>
            </div>
        </div>
        <hr />
        <button type="submit" class="btn btn-default">Export</button>
    </form:form>
</body>
</html>
