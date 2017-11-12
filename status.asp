<!DOCTYPE html>
<!--[if lt IE 7]>      <html class="no-js lt-ie9 lt-ie8 lt-ie7"> <![endif]-->
<!--[if IE 7]>         <html class="no-js lt-ie9 lt-ie8"> <![endif]-->
<!--[if IE 8]>         <html class="no-js lt-ie9"> <![endif]-->
<!--[if gt IE 8]><!--> <html class="no-js"> <!--<![endif]-->
    <head>
        <meta charset="utf-8">
        
        <title>BelHard Service</title>
        <meta name="description" content="Общество с ограниченной ответственностью «БелХард Сервис» — специализированное подразделение группы компаний «БелХард», созданное в 1995 году для сервисного обслуживания компьютерного и периферийного оборудования компании Hewlett-Packard (HP).">
				<meta name="keywords" content="ремонт HP, гарантийный ремонт, негарантийный ремонт, ремонт, Сервисный центр HP, авторизованный сервисный центр, ремонт серверов HP, Сервисный центр Hewlett-Packard, официальный сервисный центр Hewlett-Packard, официальный сервисный центр HP, ремонт ноутбуков, ремонт ноутбуков HP, ремонт принтеров, ремонт принтеров HP, ремонт ноутбука HP, ремонт принтера HP, ременонт сервера HP">
        <meta name="viewport" content="width=device-width, initial-scale=1">

         <!-- Кинуть favicon.ico и apple-touch-icon.png в корень сайта -->

        <link rel="stylesheet" href="css/bootstrap.min.css">
        <link rel="stylesheet" href="css/animations.css">
        <link rel="stylesheet" href="css/font-awesome.min.css"> 
        <link rel="stylesheet" href="css/main.css">
		<link rel="stylesheet" media="screen" href="css/styles.css" >
        <script src="js/vendor/modernizr-2.6.2.min.js"></script>
        <!--[if lt IE 9]>
            <script src="js/vendor/respond.min.js"></script>
        <![endif]-->
    </head>
    <body>
        <!--[if lt IE 7]>
             <p class="browsehappy">Вы используете <strong>старую версию</strong> браузера. Пожалуйста <a href="http://windows.microsoft.com">обновите ваш браузер</a>.</p>
        <![endif]-->
    <header id="header" class="light_section">
        <div class="container">
            <div class="row">
                <a class="navbar-brand" href="./index.asp"><img src="images/logo.png" alt=""></a>
                <div class="col-sm-12 mainmenu_wrap mobilephone mobilelogo">
                    <div class="main-menu-icon visible-xs">
                        <span></span>
                        <span></span>
                        <span></span>
                    </div>
                    <nav>
                        <ul id="mainmenu" class="menu sf-menu responsive-menu superfish">
                        <li style="margin-right:100px" class = "dim">Тел/факс: <b>(017) 203-89-85</b><li>							
                            <li class="">
                                <a href="./index.asp">ГЛАВНАЯ</a>
                            </li>
                            <li class="active">
                                <a href="./status.asp">СТАТУС ЗАКАЗА</a>
                             </li>  
                            <li class="">
                                <a href="./contact.asp">КОНТАКТЫ</a>
								 <li class="">
                                <a href="./company.asp">О КОМПАНИИ</a>
                            </li>
                            </li>
                        </ul>
                    </nav>
                </div>
            </div>
        </div>
    </header>



    
<section class="">
  <div class="container text-center">
    <div class="row">
        <h2 class="block-header">
Предлагаем Вам онлайн систему просмотра информации о сданном в ремонт оборудовании. </h2>
 

<%
order = request.form("order")
sn = request.form("sn")
sn = UCase(sn)



If order <> "" Then
strDSN = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=C:\Corporate Hosting Facility\www.service.belhard.com\data\Inter_resulting_table_97.mdb"
Set db = Server.CreateObject	("ADODB.Connection")
db.Open strDSN

strsql = "select * from Internet_source where Order_number =" & order
set RecSet = db.execute(strsql)
	If (recset.bof) and (recset.eof) Then
	error = "Такого заказа нет! Проверьте правильность введённого номера заказа!"
	db.Close
	Set db = Nothing

	ElseIf recset.fields("Serial_number") <> sn Then
	error = "ОШИБКА! Неправильно введен серийный номер!!!"
	db.Close
	Set db = Nothing

	Else



    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    pass="C:\Corporate Hosting Facility\www.service.belhard.com\log\"

    Dim fs, f, ts, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(pass&"counter_stat.txt")
    Set ts = f.OpenAsTextStream(ForReading, TristateUseDefault)
    s = ts.ReadLine
    num =CInt(s)+1
    ts.Close
    Set ts = f.OpenAsTextStream(ForWriting, TristateUseDefault)
    ts.Write ""&num
    ts.Close
'Set ts = f.OpenAsTextStream(ForAppending, TristateUseDefault)
'ts.Write ""&num&" "&order
'ts.Close

		response.write "<br><br><table width=80% cellspacing=1 cellpadding=0 border=0 align=center><tr>"
		response.write "<td class=param2 colspan=3><b>&nbsp;<h2>Просмотр информации по заказу:<h2> " & order & "</b></font></td></tr>"


		If recset.fields("Product_name") <> vbNull Then
		response.write "<tr><td class=param1 width=50% bgcolor=#dddddd><h4><b>Наименование оборудования:</b></h4></td><td class=param1 colspan=2 width=50% bgcolor=#eeeeee><h2>" & recset.fields("Product_Name").value & "</h2></td></tr>"
		End if
		If recset.fields("Received_Date") <> vbNull Then
		response.write "<tr><td class=param1 width=50% bgcolor=#dddddd><h4><b>Дата приёмки в ремонт:</b></h4></td><td class=param1 colspan=2 width=50% bgcolor=#eeeeee><h2>" & recset.fields("Received_Date").value & "</h2></td></tr>"
		End if

		If recset.fields("Started_Date") <> vbNull Then
		response.write "<tr><td class=param1 width=50% bgcolor=#dddddd><h4>Дата начала ремонта:</h4></td><td class=param1 colspan=2 width=50% bgcolor=#eeeeee><h3>" & recset.fields("Started_Date").value & "</h3></td></tr>"
		else
'расчет оринтировочного начала ремонта
		strsql = "select Order_Number from Internet_source where E_SN = 'N/A' order by Order_Number"
		set RecSet = db.execute(strsql)
		recset.movefirst
		i=0
		do while not recset.eof
		orr = CLng(order)
		if recset.fields("Order_Number") < orr Then
		i=i+1
		End if
		recset.movenext
		loop
		orr = CLng((i/10)*8)
		if orr < 1 Then orr = 1
		response.write "<tr><td class=param1 width=100% bgcolor=#dddddd colspan=3><h4>К сожалению ремонт Вашего оборудования еще не начат. Предположительно ремонт Вашего оборудования начнётся через " & orr & " рабочих часов.</h4></td></tr>"
		response.write "</td></tr></table>"
		db.Close
		Set db = Nothing
		%>
		<font class=param1><br><a href=http://www.service.belhard.com/status.asp><h2>Просмотреть информацию по следующему заказу<h2></a>
		<%
		response.end
		End if
		
		If recset.fields("Repair_TypeS") <> vbNull and recset.fields("Started_Date") <> vbNull Then
		response.write "<tr><td class=param1 width=50% bgcolor=#dddddd><h4>Тип ремонта:</h4></td><td class=param1 colspan=2 width=50% bgcolor=#eeeeee><h4><b>" & recset.fields("Repair_TypeS").value & "</b></h4></td></tr>"
		End if

		If recset.fields("Test_Results") <> vbNull and recset.fields("Started_Date") <> vbNull Then
		Select Case recset.fields("Test_Results")
		Case "TESTED OK"
		tr = "Сделано"
		Case "NOT TESTED"
		tr = "Не сделано"
		Case ""
		tr = "В работе"
		End Select	
		response.write "<tr><td class=param1 width=50% bgcolor=#dddddd><h4>Состояние заказа:</h4></td><td class=param1 colspan=2 width=50% bgcolor=#eeeeee><h4><b>" & tr & "</b></h4></td></tr>"
		End if

		If recset.fields("Finished_Date") <> vbNull Then
		response.write "<tr><td class=param1 width=50% bgcolor=#dddddd><h4>Дата окончания ремонта:</h4></td><td class=param1 colspan=2 width=50% bgcolor=#eeeeee>" & recset.fields("Finished_Date").value & "</td></tr>"
else
		Select Case recset.fields("E_SN")
		Case "Воронович"
		ph = 321
		Case "Гапочкин"
		ph = 228
		Case "Злотников"
		ph = 220
		Case "Кормызов"
		ph = 224
		Case "Короленко"
		ph = 222
		Case "Мамоненко"
		ph = 229
		Case "Нагорный"
		ph = 226
		Case "Полонский"
		ph = 221
		Case "Савко"
		ph = 322
		Case "Урбанович"
		ph = 225
		Case "Фицнер"
		ph = 229
		Case "Чемский"
		ph = 229
		End Select	
		response.write "<tr><td class=param1 width=50% bgcolor=#dddddd><h4>Фамилия инженера, выполняющего ремонт и контактный телефон:</h4></td><td class=param1 colspan=2 width=50% bgcolor=#eeeeee><h3>" & recset.fields("E_SN").value & ", " & "(017)203-89-85" & " #" & ph & "<h3></td></tr>"





		End if

		'счет
		If recset.fields("Test_Results") = "TESTED OK" and recset.fields("Repair_TypeS") = "Платный" Then
			bill=Len(recset.fields("Bill_Number"))
			if recset.fields("Bill_Number") = "N/A" or bill = 10 Then
			bill="Формируется"
			Else
			bill="№ " & recset.fields("Bill_Number") & " на сумму " & recset.fields("Bill_Total_BRB") & " руб."
			End if
                response.write "<tr><td class=param1 width=50% bgcolor=#dddddd><h4><b>Счёт:</b></h4></td><td class=param1 width=50% bgcolor=#eeeeee><h3>" & bill & "</h3></td></tr>"
'response.write "<tr><td class=param1 width=50% bgcolor=#dddddd>Фамилия бухгалтера и контактный телефон</td><td class=param1 width=50% bgcolor=#eeeeee>Варивончик (017)226-84-26 #223</td></tr>"

		End if

		response.write "</td></tr></table>"
		db.Close
		Set db = Nothing
		%>
		<font class=param1><br><a href=http://www.service.belhard.com/status.asp>Просмотреть информацию по следующему заказу</a>
		<%
		response.end
	End if
End if

%>

<SCRIPT LANGUAGE="JavaScript">
function validator(theForm) {
if (theForm.order.value == "") { alert('Вы не ввели номер заказа');
return false; }
if (theForm.sn.value == "") { alert('Вы не ввели серийный номер оборудования');
return false; }
return true;
} 
</SCRIPT>

<br><br>
<table border="0" width="100%">
<tr><td valign=top>
<font class="param1"><p align="justify">


<% response.write "<center><b>" & error & "</b>" %>
<form class="contact_form" action="status.asp" method="post" name="status" onSubmit="return validator(this)">
    <ul>
        <li>
             <h2>Статус заказа</h2>
             <span class="required_notification hiddenform">* Вы должны заполнить все поля</span>
 <span class="required_notification"><a class="hideinstruction" href="#openModal">Инструкция</a></span>

<div id="openModal" class="modalDialog">
	<div>
		<a href="#close" title="Закрыть" class="close">X</a>
		<h2>Инструкция</h2>
		<p>Для просмотра информации необходимо ввести номер заказа, который указан на расписке. <b>Ввод номера заказа должен осуществляться без дефиса.</b>
<img src=".\images\order_number.gif" border="0"><br></p>
		<p>Необходимый серийный номер также указан на расписке в разделе "Информация об оборудовании"<br><br>
<img src=".\images\s_n.gif" border="0"></p>
	</div>
</div>
        </li>	
        <li>
            <label class="sizenameandsn" for="name">Номер заказа:</label>
            <input type="text" maxlength="5" pattern="[0-9]{5}" name="order" value="<% response.write order %>" required />
			<span class="form_hint">Пример: "71345"</span>
        </li>
        <li>
            <label class="sizenameandsn" for="name">Серийный номер:</label>
            <input type="text" maxlength="10" name="sn" value="<% response.write sn %>" required />
            <span class="form_hint">Пример: "SNG3DCS5SA"</span>
        </li> 
        <li>
        	<button class="submit" type="submit" value="Отправить">Отправить</button>
        </li>
    </ul>
</td>
</tr></table>
</div>
    </div>
  </div>
 </form>
 
 
</form>
 
</section>
<section class="banner-box color_section parallax" style="margin-top:120px; padding-top:50px; padding-bottom: 10px;">
    <div class="container">
        <div class="row">
            <div class="col-sm-12 text-center">
                <p class="title">Выезд специалиста</p>
                <p>ООО "БелХард Сервис" предлагает платную услугу по выезду инженера  к Заказчику (действует в пределах Минска). Инженеры нашей компании готовы выехать на Ваш объект для консультации и решения задач в области ремонта техники HP. Предложение действительно для юридических лиц.
                </p>
            </div>
        </div>
    </div>
	<br><br><br>
</section>
<footer id="header" class="color_section" style="padding-top: 90px;
padding-bottom: 19px; padding-top:15px; margin-bottom:0;">
        <div class="container">
            <div class="row">

                <div class="col-sm-12 text-center">
                    &copy; Copyright 1999-2017 <br>
              		<a href="mailto:service@belhard.com">BelHard Service</a>
                </div>
            </div>

        </div>
    </footer>
 

    <div class="preloader">
        <div class="preloaderimg"></div>
    </div>
        <script src="js/vendor/jquery-1.11.1.min.js"></script>
        <script src="js/vendor/jquery-migrate-1.2.1.min.js"></script>
        <script src="js/vendor/bootstrap.min.js"></script>
        <script src="js/vendor/placeholdem.min.js"></script>
        <script src="js/vendor/hoverIntent.js"></script>
        <script src="js/vendor/superfish.js"></script>
        <script src="js/vendor/jquery.actual.min.js"></script>
        <script src="js/vendor/jquery.appear.js"></script>
        <script src="js/vendor/jquerypp.custom.js"></script>
        <script src="js/vendor/jquery.elastislide.js"></script>
        <script src="js/vendor/jquery.flexslider-min.js"></script>
        <script src="js/vendor/jquery.prettyPhoto.js"></script>
        <script src="js/vendor/jquery.easing.1.3.js"></script>
        <script src="js/vendor/jquery.ui.totop.js"></script>
        <script src="js/vendor/jquery.isotope.min.js"></script>
        <script src="js/vendor/jquery.easypiechart.min.js"></script>
        <script src='js/vendor/jflickrfeed.min.js'></script>
        <script src="js/vendor/jquery.sticky.js"></script>
        <script src='js/vendor/owl.carousel.min.js'></script>
        <script src='js/vendor/jquery.nicescroll.min.js'></script>
        <script src='js/vendor/jquery.fractionslider.min.js'></script>
        <script src='js/vendor/jquery.scrollTo-min.js'></script>
        <script src='js/vendor/jquery.localscroll-min.js'></script>
        <script src='js/vendor/jquery.parallax-1.1.3.js'></script>
        <script src='js/vendor/jquery.bxslider.min.js'></script>
        <script src='js/vendor/jquery.funnyText.min.js'></script>
        <script src='js/vendor/jquery.countTo.js'></script>
        <script src="js/vendor/grid.js"></script>
        <script src='twitter/jquery.tweet.min.js'></script>
        <script src="js/plugins.js"></script>
        <script src="js/main.js"></script>
    </body>
</html>