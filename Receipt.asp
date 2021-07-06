<% @ Language="JavaScript" %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8" />
    <title></title>
</head>
<body>

    <p>Find Tickets.com</p>
    <p>Your booking has been comfirmed!</p>

    <p>Your Booking Receipt</p>
    Name:
    <%
    var fname
    fname=Request.Form("bname")

    Response.Write(fname)
    %>

    <br />

    <br />
    Movie:

    <%
    var moname
    moname=Request.Form("mname")

    Response.Write(moname)

    %>

    <br />
    <br />

    Date:
    <%
    var bookdate
    bookdate=Request.Form("bdate")

    Response.Write(bookdate)



    %>
    <br />
    <br />
    Booking Quantity :
    <%
    var bquantity
    bquantity=Request.Form("quantity")

    Response.Write(bquantity)

    %>

    <br />
    <br />
    Reference number:
    <%
    var  ref = "Findticket2021" + Math.floor(Math.random() * (10000 - 1) + 1);
    Response.Write(ref)

    %>
    <br />

    <%
    var conn = new ActiveXObject("ADODB.Connection");

        conn.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source='‪D:/Bookingdatabase.mdb'");

    SqlString = "INSERT INTO Bookings(BookingName, MovieName, BookingDate, TicketQuantity, ReferenceID) VALUES('" + fname + "', '" + moname + "', '" + bookdate + "','" + bquantity + "','" + ref + "')";
    conn.Execute(SqlString);
    %>
</body>
</html>