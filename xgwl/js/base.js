// 选项卡 鼠标经过切换
$(".TAB li").mousemove(function () {
    var tab = $(this).parent(".TAB");
    var con = tab.attr("id");
    var on = tab.find("li").index(this);
    $(this).addClass('on').siblings(tab.find("li")).removeClass('on');
    $(con).eq(on).show().siblings(con).hide();
});
//打开弹窗
function openfrom(v){

        $("#downurl").val(v);
        $(".intpop .wrtit").html("请填写表格");
        $(".intpop").fadeIn();
        $(".mask").fadeIn();
        console.log(v);

}
function openfrom2(v){
    if($.cookie('downcheck')=="-_-||"){
        export_raw(v);
        console.log($.cookie('downcheck'));
    }else{
        $("#downurl").val(v);
        $(".intpop .wrtit").html("请填写表格，获取资料");
        $(".intpop").fadeIn();
        $(".mask").fadeIn();
        console.log(v);
    }
}

// 选项卡 鼠标点击
$(".TAB_CLICK li").click(function () {
    var tab = $(this).parent(".TAB_CLICK");
    var con = tab.attr("id");
    var on = tab.find("li").index(this);
    $(this).addClass('on').siblings(tab.find("li")).removeClass('on');
    $(con).eq(on).show().siblings(con).hide();
});
// 选项卡 鼠标点击
$(".TAB_CLICKa a").click(function () {
    var tab = $(this).parent(".TAB_CLICKa");
    var con = tab.attr("id");
    var on = tab.find("a").index(this);
    $(this).addClass('on').siblings(tab.find("a")).removeClass('on');
    $(con).eq(on).show().siblings(con).hide();
});

// pc端导航切换
$(".XGnav .XGli").mousemove(function () {
    var tab = $(this).parent(".XGnav");
    var con = tab.attr("id");
    var on = tab.find(".XGli").index(this);
    $(this).addClass('on').siblings(tab.find(".XGli")).removeClass('on');
    if (on != 0 && on != 1) {
        $(con).eq(on).fadeIn().siblings(con).fadeOut();
    } else {
        $(con).fadeOut();
    }
});
$(document).click(function () {
    $(".pc-nav").fadeOut();
});
$(".pc-nav").mouseleave(function () {
    //  $(".pc-nav").fadeOut();
});
//向下跳转锚
$("a.topLink").click(function () {
    $("html, body").animate({
        scrollTop: $($(this).attr("href")).offset().top + "px"
    }, {
        duration: 1000,
        easing: "swing"
    });
    return false;
});

//显示弹窗并定时关闭弹窗
function showpop() {
    $('.pop').fadeIn();
    setTimeout("$('.pop').fadeOut()", 2000);

}
$(".cpli").click(function () {
    showpop();
});
$('.pop').click(function () {
    $('.pop').fadeOut();
});
$(document).ready(function () {
    console.log("最后更新日期2017.3.18 by 任刚");
});
//关闭电脑端默认二级菜单
if (!navigator.userAgent.match(/(iPhone|iPod|iPad|Android|ios)/i)) {
    $(document).off('click.bs.dropdown.data-api');
}
;
//在线咨询
function feedback() {
    alert("在线咨询弹窗")

}
$(".sxf").click(function(){
   var aaa=$(this).parents(".searchint").find(".sint").val();

    location='4.asp?key='+aaa;
})
$(".s2sf").click(function(){
    var aaa=$(this).parents(".searchint").find(".sint").val();

   location='3.asp?key='+aaa;
})

//input
