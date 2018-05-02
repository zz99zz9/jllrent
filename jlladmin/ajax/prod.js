$(".save").click(function(){

update();
});
/////////////
function update(){
classid=$("#classid").val();

    prodname=$("#prod_name").val();
	uptime=$("#uptime").val();
    DefaultPicUrl=$("#DefaultPicUrl").val();
    content=$("#content").val();
	Oid=$("#Oid").val();
/*	BigClassName=$("#BigClass").val().split(";")[0];
 BigClassId=$("#BigClass").val().split(";")[1];
 SmallClassName=$("#SmallClass").val().split(";")[0];
 SmallClassId=$("#SmallClass").val().split(";")[1];*/
  console.log(SmallClassId);
	Keywords=$("#Keywords").val();
	Description=$("#Description").val();
	Passed=$("#Passed").is(':checked')
	Elite=$("#Elite").is(':checked')
	action=$("#action").val();//操作指令
	pid=$("#pid").val();//产品系统编号
	//后加值
tips=$("#tips").val();//产品系统编号
picalt=$("#PicAlt").val();

console.log(action);
	/*	  BigClassName:BigClassName,
	 BigClassId:BigClassId,
	 SmallClassName:SmallClassName,
	 SmallClassId:SmallClassId,*/
  $.post("ajax/prod.asp",{
	  classid:classid,
	  prodname:prodname,
	  DefaultPicUrl:DefaultPicUrl,
      content:content,
	  Oid:Oid,
	  uptime:uptime,
	  picalt:picalt,
	  Description:Description,
	  Keywords:Keywords,
      Passed:Passed,
	  Elite:Elite,
	  action:action,
	  pid:pid,
	  tips:tips
	  },
	  
    function(data,status){
	  if(status=="success"){
	//	  console.log(data);
		   if(data=="add"){

		  location.href='Prod_List.Asp?sta=add&Prod_Type='+classid+''
		  }else if(data=="mod"){

			  location.href='Prod_List.Asp?sta=mod&Prod_Type='+classid+''
			  }else if(data=="del"){

			  location.href='Prod_List.Asp?sta=del&Prod_Type='+classid+''
		  }
		/* */

		  }
    });
}