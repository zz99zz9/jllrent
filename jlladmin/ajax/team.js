$(".save").click(function(){

update();
});
/////////////
function update(){
classid=$("#classid").val();
id=$("#id").val();
    cname=$("#cname").val();
	ename=$("#ename").val();
	uptime=$("#uptime").val();
    DefaultPicUrl=$("#pic").val();
	position=$("#position").val();
	Oid=$("#oid").val();
	tel=$("#tel").val();
	summary=$("#summary").val();
	mail=$("#mail").val();
	info=$("#info").val();
	action=$("#action").val();//操作指令


console.log(classid);

  $.post("ajax/team.asp",{
	  classid:classid,
	  id:id,
	  cname:cname,
	  ename:ename,
	  DefaultPicUrl:DefaultPicUrl,
      position:position,
	  Oid:Oid,
	  uptime:uptime,
	  tel:tel,
	  summary:summary,
	  mail:mail,
      info:info,
	  action:action
	  },
    function(data,status){
	  if(status=="success"){
	//	  console.log(data);
		   if(data=="add"){

		  location.href='team_List.Asp?sta=add&classid='+classid+''
		  }else if(data=="mod"){

			  location.href='team_List.Asp?sta=mod&classid='+classid+''
			  }else if(data=="del"){

			  location.href='team_List.Asp?sta=del&classid='+classid+''
		  }
		/* */

		  }
    });
}