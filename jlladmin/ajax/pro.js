$(".save").click(function(){
console.log("O_O");
update();
});
$(".save2").click(function(){

update("cg");
});
/////////////
function update(cg){
	  BigClassName=$("#BigClass").val().split(";")[0];
  BigClassId=$("#BigClass").val().split(";")[1];
  SmallClassName=$("#SmallClass").val().split(";")[0];
  SmallClassId=$("#SmallClass").val().split(";")[1];
  Product_Id=$("#Product_Id").val();
  Title=$("#Title").val();
  uptime=$("#uptime").val();
	gid=$("#gid").val();
	console.log(gid);
  DefaultPicUrl=$("#DefaultPicUrl").val();
  DefaultPicUrl2=$("#DefaultPicUrl2").val();
    DefaultPicUrl3=$("#DefaultPicUrl3").val();
    content=$("#content").val();
	Orderid=$("#Orderid").val();
	picalt=$("#picalt").val();
	  if(cg=="cg"){
  Passed=$("#Passed").is(!this.checked);
  }else{
	 Passed=$("#Passed").is(':checked') ;
  }
	Elite=$("#Elite").is(':checked');
	dl=$("#dl").is(':checked');
	console.log(dl);
	action=$("#action").val();//操作指令
	pro_id=$("#pro_id").val();//产品系统编号
	//后加值
	enTitle=$("#enTitle").val();
	 
	Keywords=$("#Keywords").val();
	Description=$("#Description").val();
	jgzj=$("#jgzj").val();
	jgjj=$("#jgjj").val();
	mj1=$("#mj1").val();
	mj2=$("#mj2").val();
	cfjname=$("#cfj").val().split(";")[0];
	cfjid=$("#cfj").val().split(";")[1];
/*
	clxname=$("#clx").val().split(";")[0];
	clxid=$("#clx").val().split(";")[1];
	*/
	clbname=$("#clb").val().split(";")[0];
	clbid=$("#clb").val().split(";")[1];
	crqname=$("#crq").val().split(";")[0];
	crqid=$("#crq").val().split(";")[1];
	content1=$("#content1").val();
	content2=$("#content2").val();
	//二次新加
	ckfsname=$("#ckfs").val().split(";")[0];
	ckfsid=$("#ckfs").val().split(";")[1];
	wz=$("#wz").val();
	syq=$("#syq").val();
	jgsj=$("#jgsj").val();
	qt=$("#qt").val();
	hx=$("#hx").val();

//jquery获取复选框值      
            var ctdid =[];//定义一个数组      
            $('input[name="ctd"]:checked').each(function(){//遍历每一个名字为interest的复选框，其中选中的执行函数      
            ctdid.push($(this).val());//将选中的值添加到数组chk_value中      
            }); 
	ctdid=ctdid.toString()
var clxid =[];//定义一个数组      
            $('input[name="clx"]:checked').each(function(){//遍历每一个名字为interest的复选框，其中选中的执行函数      
            clxid.push($(this).val());//将选中的值添加到数组chk_value中      
            }); 
	clxid=clxid.toString()
	
//	console.log(ctdid);
//	console.log(jgjj);

  $.post("ajax/pro.asp",{
	  BigClassName:BigClassName,
	  BigClassId:BigClassId,
	  SmallClassName:SmallClassName,
	  SmallClassId:SmallClassId,
	  Product_Id:Product_Id,
	  picalt:picalt,
	  Title:Title,
	  gid:gid,
	  Description:Description,
	  Keywords:Keywords,
	  uptime:uptime,
	  DefaultPicUrl:DefaultPicUrl,
	  DefaultPicUrl2:DefaultPicUrl2,
	  DefaultPicUrl3:DefaultPicUrl3,
      content:content,
	  Orderid:Orderid,
      Passed:Passed,
	  Elite:Elite,
	  dl:dl,
	  action:action,
	  pro_id:pro_id,
	  //新增
	  enTitle:enTitle,
	  content1:content1,
	  content2:content2,
	  jgzj:jgzj,
	  jgjj:jgjj,
	  mj1:mj1,
	  mj2:mj2,
	  cfjname:cfjname,
	  cfjid:cfjid,
	  crqname:crqname,
	  crqid:crqid,
//	  clxname:clxname,
	  clxid:clxid,
	  clbname:clbname,
	  clbid:clbid,
//	  ctdname:ctdname,
	  ctdid:ctdid,
	  //二次新增
	ckfsname:ckfsname,
	ckfsid:ckfsid,
	wz:wz,
	syq:syq,
	jgsj:jgsj,
	qt:qt,
	hx:hx},
	  
    function(data,status){
	  if(status=="success"){
		//  console.log(data);
		  /**/
		 if(data=="add"){
 					location.href='pic_list.Asp'
			}else if(data=="addcg"){
					location.href='pro_mod.Asp'
			}else if(data=="modcg"){
					location.href='pro_mod.Asp'
		 	 }else if(data=="mod"){
				  location.href='Pro_List.Asp?sta=mod'
			 }else if(data=="del"){
				  location.href='Pro_List.Asp?sta=del'
		  }
		/* */

		  }
    });
}