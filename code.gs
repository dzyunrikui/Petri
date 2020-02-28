function emailsend(user_num,name,email ) 
{
 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var d = new Date();
  d.setHours(0);
  d.setMinutes(0);
  d.setSeconds(0);
  d.setMilliseconds(0);

  for (var i=1; i<= sheet.getLastRow(); i++) //krutim rows
  {
          { 
            var data = sheet.getRange(i,user_num).getValue();  // v perem data kladem znechine (i,j) yacheiki
             
            if (data.getTime && data.getTime() === d.getTime())
            {  if (sheet.getRange(i, user_num).getBackground()=== "#ff0000")
              {}
             else  if (sheet.getRange(i, user_num).getBackground()=== "#ffff0000")
              {}
                else  if (sheet.getRange(i, user_num).getBackground()=== "#00ff00")
              {}
               else  if (sheet.getRange(i, user_num).getBackground()=== "#ff00ff00")
              {}
               else  if (sheet.getRange(i, user_num).getBackground()=== "#9900ff")
              {}
               else  if (sheet.getRange(i, user_num).getBackground()=== "#ff9900ff")
              {}
             else
             { var subject = "Перекрась ячейку";1
                  var message = "Привет " + name +"!"+"  Ты играешь  "+ sheet.getRange(i, 1).getValue()+ "?"+ "  Если играешь  " + subject + " и если не играешь все равно перекрась :))  https://docs.google.com/spreadsheets/d/1TO-3_rnKYhDQ0cjthJG9SJbkGKVFpQt5rk2ew1lZuVE/edit#gid=0 ";   
                  MailApp.sendEmail(email, subject, message);
              }
           }
  }
} 
}



function check_color(user_num,name,email ) 
{
 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var d = new Date();
  var subject = "Сегодня игра";
  d.setHours(0);
  d.setMinutes(0);
  d.setSeconds(0);
  d.setMilliseconds(0);
  
  for (var matchday=1; matchday<= sheet.getLastRow(); matchday++) //ishchem datu igry
     {
            var data = sheet.getRange(matchday,3).getValue();  // v perem data kladem datu igry
                
            if (data.getTime && data.getTime() === d.getTime())
             {
             var players=0
             for (var j=6; j<=16; j++)
             {  
             // kol-vo igrokov
            if (sheet.getRange(matchday, j).getBackground() === "#00ff00" ) 
                 { players=players+1}
               else if (sheet.getRange(matchday, j).getBackground() === "#ff00ff00" )
                {players=players+1}
               } 
               if  (players>=3){
          //  sheet.getRange(matchday, user_num).setValue(sheet.getRange(matchday, user_num).getBackground())       
           if (sheet.getRange(matchday, user_num).getBackground() === "#00ff00")    
      
               {    
                if (sheet.getRange(matchday, 5).isBlank()) { var dengi="не понятно"}
                 else {var dengi=sheet.getRange(matchday, 5).getValue()/players}
                  var message = "Привет " + name +"!"+ "  Напоминаю, сегодня ты играешь   "+sheet.getRange(matchday, 1).getValue()+" на площадке "+ sheet.getRange(matchday, 2).getValue()+ ". Игра начнется в "+ sheet.getRange(matchday, 4).getValue()+"  С тебя  "+ dengi+ " рублей"  ;   
             //     sheet.getRange(matchday, user_num).setValue(email)
                  MailApp.sendEmail(email, subject, message);
                 }             
               else if (sheet.getRange(matchday, user_num).getBackground()=== "#ff00ff00")
               {    
                if (sheet.getRange(matchday, 5).isBlank()) { var dengi="не понятно"}
                 else {var dengi=sheet.getRange(matchday, 5).getValue()/players}
                  var message = "Привет " + name +"!"+ "  Напоминаю, сегодня ты играешь   "+sheet.getRange(matchday, 1).getValue()+" на площадке "+ sheet.getRange(matchday, 2).getValue()+ ". Игра начнется в "+ sheet.getRange(matchday, 4).getValue()+"  С тебя  "+ dengi+ " рублей"  ;   
                  MailApp.sendEmail(email, subject, message);
                 }     
                 
               }
              }   
      }
}

function pinok(user_num ) 
{ var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var d = new Date()
  d.setHours(0);
  d.setMinutes(0);
  d.setSeconds(0);
  d.setMilliseconds(0);
 for (var matchday=1; matchday<= sheet.getLastRow(); matchday++) //ishchem datu igry
   {
        var data = sheet.getRange(matchday,3).getValue();
        diff = (data-d)/(1000*3600*24)
        if (diff ==4)
          {
           if (sheet.getRange(matchday, user_num).getBackground() === "#ffffff")
           
           { var dd=new Date(new Date().getTime() + 10*24 * 60 * 60 * 100)
                dd.setHours(0);
                dd.setMinutes(0);
                dd.setSeconds(0);
                dd.setMilliseconds(0);
                sheet.getRange(matchday, user_num).setValue(dd)
              }
          }
   }
}
function dokraska() 
{ var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  for (var matchday=1; matchday<= sheet.getLastRow(); matchday++) //ishchem datu igry
{ 
  var players=0
   for (var j=6; j<=14; j++)  // kol-vo igrokov
   {          
            if (sheet.getRange(matchday, j).getBackground() === "#ff0000") 
                 { players=players+1}
   }
  if (players>=4)
  { 
    if (players<=9)
        {for (var k=6; k<=16; k++ )
    //sheet.getRange(matchday, k).setBackground("#ff0000")
    sheet.getRange(matchday, k).setValue(sheet.getRange(matchday, k).getBackground())
    
    }
  }
}
} 

function vacation(user_num ) 
{ var ss = SpreadsheetApp.getActiveSpreadsheet();
  var main_sheet = ss.getSheets()[0];
  var email_sheet = ss.getSheets()[2];
  var vacation=email_sheet.getRange(user_num, 2).getValue();  
  for (var matchday=3; matchday<= main_sheet.getLastRow(); matchday++) //ishchem datu igry
   
  if (main_sheet.getRange(matchday,3).isBlank())  {}
  else 
   if (main_sheet.getRange(matchday, user_num).getBackground() ==="#ff0000") {}
              else  if (main_sheet.getRange(matchday, user_num).getBackground()=== "#ffff0000")
              {}
                else  if (main_sheet.getRange(matchday, user_num).getBackground()=== "#00ff00")
              {}
               else  if (main_sheet.getRange(matchday, user_num).getBackground()=== "#ff00ff00")
              {}
               else  if (main_sheet.getRange(matchday, user_num).getBackground()=== "#9900ff")
              {}
               else  if (main_sheet.getRange(matchday, user_num).getBackground()=== "#ff9900ff")
              {}
   else
   {
    var data = main_sheet.getRange(matchday,3).getValue();
    diff = (data-vacation)/(1000*3600*24)
    if (diff <0)
    {   main_sheet.getRange(matchday, user_num).setBackground("#ff0000")
        main_sheet.getRange(matchday, user_num).setValue("Отпуск")} 
     }
  
}

function Outdated_games_eraser() 
{ 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var d = new Date()
  d.setHours(0);
  d.setMinutes(0);
  d.setSeconds(0);
  d.setMilliseconds(0);
  for (var matchday=3; matchday< sheet.getLastRow(); matchday++) //ishchem datu igry
   {
    var data = sheet.getRange(matchday,3).getValue();
    diff = (data-d)/(1000*3600*24)
    if (diff <=-1)
    {   sheet.deleteRow(matchday)} 
  
   }

}


function alexandr ()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
//   sheet.getRange(15, 2).setValue('привет') // лист с почтами 
  var user_num=6 ; 
  if (sheet.getRange(user_num,3).isBlank())  { }
  else
  {
  emailsend(user_num,sheet.getRange(user_num,1).getValue(),sheet.getRange(user_num,3).getValue())
  check_color(user_num,sheet.getRange(user_num,1).getValue(),sheet.getRange(user_num,3).getValue())
  pinok(user_num)
  vacation(user_num) 
  }
  
}



function denis ()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var user_num=7  ;
  if (sheet.getRange(user_num,3).isBlank())   {}
  else
  {
  emailsend(user_num,sheet.getRange( user_num,1).getValue(),sheet.getRange(user_num,3).getValue())
  check_color(user_num,sheet.getRange( user_num,1).getValue(),sheet.getRange(user_num,3).getValue())
  pinok(user_num)
  vacation(user_num) 
  }
}
function leonid ()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var user_num=8  ;
  if (sheet.getRange(user_num,3).isBlank())   {}
  else
  {
  emailsend(user_num,sheet.getRange(user_num,1).getValue(),sheet.getRange(user_num,3).getValue())
  check_color(user_num,sheet.getRange( user_num,1).getValue(),sheet.getRange(user_num,3).getValue())
  pinok(user_num)
  vacation(user_num) 
  }
}
function lera ()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var user_num=9  ;
  if (sheet.getRange(user_num,3).isBlank())   {}
  else
  {
  emailsend(user_num,sheet.getRange( user_num,1).getValue(),sheet.getRange(user_num,3).getValue())
  check_color(user_num,sheet.getRange( user_num,1).getValue(),sheet.getRange(user_num,3).getValue())
  pinok(user_num)
  vacation(user_num) 
  }
}

function maxim ()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var user_num=10  ;
  if (sheet.getRange(user_num,3).isBlank())  {}
  else
  {
  emailsend(user_num,sheet.getRange( user_num,1).getValue(),sheet.getRange(user_num,3).getValue())
  check_color(user_num,sheet.getRange( user_num,1).getValue(),sheet.getRange(user_num,3).getValue())
  pinok(user_num)
  vacation(user_num) 
  }
}


function marina ()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var user_num=11  ;
  if (sheet.getRange(user_num,3).isBlank())  {}
  else
  {
  emailsend(user_num,sheet.getRange( user_num,1).getValue(),sheet.getRange(user_num,3).getValue())
  check_color(user_num,sheet.getRange( user_num,1).getValue(),sheet.getRange(user_num,3).getValue())
  pinok(user_num)
  vacation(user_num) 
  }
 }


function nikolai ()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var user_num=12 ; 
  if (sheet.getRange(user_num,3).isBlank())  {}
  else
  {
  emailsend(user_num,sheet.getRange( user_num,1).getValue(),sheet.getRange(user_num,3).getValue())
  check_color(user_num,sheet.getRange( user_num,1).getValue(),sheet.getRange(user_num,3).getValue())
  pinok(user_num)
  vacation(user_num) 
  }
}


function julia ()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var user_num=13 ; 
  if (sheet.getRange(user_num,3).isBlank())   {}
  else
  {
  emailsend(user_num,sheet.getRange( user_num,1).getValue(),sheet.getRange(user_num,3).getValue())
  check_color(user_num,sheet.getRange( user_num,1).getValue(),sheet.getRange(user_num,3).getValue())
  pinok(user_num)
  vacation(user_num) 
  }
}

