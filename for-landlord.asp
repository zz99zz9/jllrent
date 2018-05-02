<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Config.Asp"-->
<!--#include file="inc/Function.asp"-->
<!--#include file="inc/Inc.Asp"-->
<%
dim tdkid
tdkid=2
%>
<!--#include file="./inc/tdk.asp"-->
<%dim bc,sc,fj,lx,lb,order,page
key=request.QueryString("key")
bc=request.QueryString("bc")
sc=request.QueryString("sc")
fj=request.QueryString("fj")
lx=request.QueryString("lx")
lb=request.QueryString("lb")
order=request.QueryString("order")
page=request.QueryString("page")
if bc="" then bc=0
if sc="" then sc=0
if fj="" then fj=0
if lx="" then lx=0
if lb="" then lb=0
if page="" then page=1
if order="" then order=0

%>
<!--#include file="inc/header.asp"-->
    <link rel="stylesheet" href="xgwl/css/5.css"/>
        <style>
        .ninfo .ntxt{border-top:0px;}
        .part1{margin-top:140px;}
        h1.eloqua-heading, h2.eloqua-heading {
    font-size: 24px;
}
.webform-component-fieldset.row .panel-heading .panel-title {
    font-size: 14px;
    color: #333;
    letter-spacing: 0;
    line-height: 35px;
}
.webform-component-select .select2-container--default .select2-selection--single .select2-selection__rendered {
    line-height: initial;
    font-size: 13px;
}
.olm-btn {
    border-radius: 2px;
    font-size: 1.5rem;
    height: 50px;
    line-height: 50px;
    margin-top: 2rem;
    padding: 0;
    vertical-align: middle;
}
.olm-btn-small {
    width: 250px;
}
.element-invisible {
    position: absolute !important;
    clip: rect(1px 1px 1px 1px);
    clip: rect(1px,1px,1px,1px);
    overflow: hidden;
    height: 1px;
}
.form-control {
    width: 100%;
    height: 34px;
    padding: 6px 12px;
    background-color: #fff;
    border: 1px solid #ccc;
    border-radius: 4px;
    box-shadow: inset 0 1px 1px rgba(0,0,0,.075);
    transition: border-color ease-in-out .15s,box-shadow ease-in-out .15s;
}
.Form__input {
    width: 100%;
    outline: 0;
    border: 1px solid #979797;
    border-radius: 2px;
    padding: 10px 11px;
    color: #333;
    letter-spacing: 0;
    font-size: 13px;
    height: 40px;
    line-height: initial;
}
.form-textarea-wrapper textarea {
    display: block;
    margin: 0;
    width: 100%;
    -moz-box-sizing: border-box;
    -webkit-box-sizing: border-box;
    box-sizing: border-box;
}
textarea.form-control {
    height: auto;
}
textarea.Form__input {
    min-height: 110px;
}
.panel-default > .panel-heading {
    color: #333;
    background-color: #fff;
    border-color: #ddd;
}
.panel-default {
    border-color: #fff;
}
.panel-body{padding:0;}
.form-actions{margin-bottom:50px;}

    </style>
<!--广告部份-->

<!--广告部份-->
<div class="part1 ">
        <div class="container container2">
          
            <div class="row">
                    <article id="node-212" class="node node-webform clearfix" about="/office-leasing/for-landlord" typeof="sioc:Item foaf:Document">
                        <header>
                                <span property="dc:title" content="投放房源" class="rdf-meta element-hidden"></span>      </header>
                        <div class="container container2"><form class="webform-client-form webform-client-form-212" action="/office-leasing/for-landlord" method="post" id="webform-client-form-212" accept-charset="UTF-8"><div class="form-item webform-component webform-component-markup webform-component--header form-group form-item form-type-markup form-group"><div class="row">
                    <div class="col-xs-12">
                    <h1 class="eloqua-heading">想要投放房源？</h1>
                    </div>
                    </div>
                    <div class="row">
                    <div class="col-xs-12">
                    <p>想要出租您的物业？想要出售您的办公室？只需填写以下表格或者直接致电：</p>
                    <p>*表示必填项</p>
                    </div>
                    </div>
                    <div class="row">
                    <div class="col-xs-12">
                    <div class="Form__line"> </div>
                    </div>
                    </div>
                    </div><fieldset  class="webform-component-fieldset row webform-component--contact-info-row1 panel panel-default form-wrapper">
                              <legend class="panel-heading">
                          <span class="panel-title fieldset-legend">您的联系方式</span>
                        </legend>
                              <div class="panel-body">
                            <div class="col-sm-2 form-item webform-component webform-component-select webform-component--contact-info-row1--salutation form-group form-item form-item-submitted-contact-info-row1-salutation form-type-select form-group"><select class="form-control form-select" id="edit-submitted-contact-info-row1-salutation" name="submitted[contact_info_row1][salutation]"><option value="Mr." selected="selected">先生</option><option value="Ms.">女士</option></select> <label class="control-label element-invisible" for="edit-submitted-contact-info-row1-salutation">Salutation</label>
                    </div><div class="col-sm-3 form-item webform-component webform-component-textfield webform-component--contact-info-row1--first-name form-group form-item form-item-submitted-contact-info-row1-first-name form-type-textfield form-group"><input placeholder="名" class="Form__input form-control form-text" type="text" id="edit-submitted-contact-info-row1-first-name" name="submitted[contact_info_row1][first_name]" value="" size="60" maxlength="128" /> <label class="control-label element-invisible" for="edit-submitted-contact-info-row1-first-name">名</label>
                    </div><div class="col-sm-3 form-item webform-component webform-component-textfield webform-component--contact-info-row1--last-name form-group form-item form-item-submitted-contact-info-row1-last-name form-type-textfield form-group"><input required="required" placeholder="姓*" class="Form__input required form-control form-text required" type="text" id="edit-submitted-contact-info-row1-last-name" name="submitted[contact_info_row1][last_name]" value="" size="60" maxlength="128" /> <label class="control-label element-invisible" for="edit-submitted-contact-info-row1-last-name">姓 <span class="form-required" title="This field is required.">*</span></label>
                    </div><div class="col-sm-4 form-item webform-component webform-component-textfield webform-component--contact-info-row1--company form-group form-item form-item-submitted-contact-info-row1-company form-type-textfield form-group"><input placeholder="公司名称" class="Form__input form-control form-text" type="text" id="edit-submitted-contact-info-row1-company" name="submitted[contact_info_row1][company]" value="" size="60" maxlength="128" /> <label class="control-label element-invisible" for="edit-submitted-contact-info-row1-company">公司</label>
                    </div>  </div>
                      </fieldset>
                    <fieldset  class="webform-component-fieldset row webform-component--contact-info-row2 panel panel-default form-wrapper">
                          <div class="panel-body">
                            <div class="col-sm-4 form-item webform-component webform-component-textfield webform-component--contact-info-row2--mobilePhone form-group form-item form-item-submitted-contact-info-row2-mobilephone form-type-textfield form-group"><input required="required" placeholder="手机*(+86xxxxxxxxxxx)" class="Form__input required form-control form-text required" type="text" id="edit-submitted-contact-info-row2-mobilephone" name="submitted[contact_info_row2][mobilePhone]" value="" size="60" maxlength="128" /> <label class="control-label element-invisible" for="edit-submitted-contact-info-row2-mobilephone">电话 <span class="form-required" title="This field is required.">*</span></label>
                    </div><div class="col-sm-4 form-item webform-component webform-component-email webform-component--contact-info-row2--email-address form-group form-item form-item-submitted-contact-info-row2-email-address form-type-webform-email form-group"><input required="required" class="email Form__input required form-control form-text form-email required" placeholder="电邮*" type="email" id="edit-submitted-contact-info-row2-email-address" name="submitted[contact_info_row2][email_address]" size="60" /> <label class="control-label element-invisible" for="edit-submitted-contact-info-row2-email-address">电邮 <span class="form-required" title="This field is required.">*</span></label>
                    </div><div class="col-sm-4 form-item webform-component webform-component-select webform-component--contact-info-row2--city form-group form-item form-item-submitted-contact-info-row2-city form-type-select form-group"><select class="form-control form-select" id="edit-submitted-contact-info-row2-city" name="submitted[contact_info_row2][city]"><option value="" selected="selected">选择城市</option><option value="Shanghai">上海</option><option value="Beijing">北京</option><option value="Guangzhou">广州</option><option value="Chengdu">成都</option><option value="Shenzhen">深圳</option><option value="Other">其他</option></select> <label class="control-label element-invisible" for="edit-submitted-contact-info-row2-city">City</label>
                    </div>  </div>
                      </fieldset>
                    <fieldset  class="webform-component-fieldset row webform-component--building-info-row panel panel-default form-wrapper">
                              <legend class="panel-heading">
                          <span class="panel-title fieldset-legend">房屋信息</span>
                        </legend>
                              <div class="panel-body">
                            <div class="col-sm-3 form-item webform-component webform-component-textfield webform-component--building-info-row--buildingname form-group form-item form-item-submitted-building-info-row-buildingname form-type-textfield form-group"><input placeholder="社区名称" class="Form__input form-control form-text" type="text" id="edit-submitted-building-info-row-buildingname" name="submitted[building_info_row][buildingname]" value="" size="60" maxlength="128" /> <label class="control-label element-invisible" for="edit-submitted-building-info-row-buildingname">大厦名称</label>
                    </div><div class="col-sm-3 form-item webform-component webform-component-textfield webform-component--building-info-row--unit form-group form-item form-item-submitted-building-info-row-unit form-type-textfield form-group"><input placeholder="单位/楼层" class="Form__input form-control form-text" type="text" id="edit-submitted-building-info-row-unit" name="submitted[building_info_row][unit]" value="" size="60" maxlength="128" /> <label class="control-label element-invisible" for="edit-submitted-building-info-row-unit">Unit</label>
                    </div><div class="col-sm-3 form-item webform-component webform-component-textfield webform-component--building-info-row--unitsize form-group form-item form-item-submitted-building-info-row-unitsize form-type-textfield form-group"><input placeholder="面积" class="Form__input form-control form-text" type="text" id="edit-submitted-building-info-row-unitsize" name="submitted[building_info_row][unitsize]" value="" size="60" maxlength="128" /> <label class="control-label element-invisible" for="edit-submitted-building-info-row-unitsize">面积</label>
                    </div><div class="col-sm-3 form-item webform-component webform-component-textfield webform-component--building-info-row--rent form-group form-item form-item-submitted-building-info-row-rent form-type-textfield form-group"><input placeholder="预期租金/月" class="Form__input form-control form-text" type="text" id="edit-submitted-building-info-row-rent" name="submitted[building_info_row][rent]" value="" size="60" maxlength="128" /> <label class="control-label element-invisible" for="edit-submitted-building-info-row-rent">Rent</label>
                    </div>  </div>
                      </fieldset>
                    <fieldset  class="webform-component-fieldset row webform-component--your-message panel panel-default form-wrapper">
                              <legend class="panel-heading">
                          <span class="panel-title fieldset-legend">留言</span>
                        </legend>
                              <div class="panel-body">
                            <div class="col-sm-12 form-item webform-component webform-component-textarea webform-component--your-message--message form-group form-item form-item-submitted-your-message-message form-type-textarea form-group"><div class="form-textarea-wrapper"><textarea placeholder="请留言" class="Form__input form-control form-textarea" id="edit-submitted-your-message-message" name="submitted[your_message][message]" cols="60" rows="5"></textarea></div> <label class="control-label element-invisible" for="edit-submitted-your-message-message">Your Message</label>
                    </div>  </div>
                      </fieldset>
                    <fieldset  class="webform-component-fieldset row webform-component--selections panel panel-default form-wrapper">
                          <div class="panel-body">
                            <div class="col-sm-4 Form__checkbox form-item webform-component webform-component-checkboxes webform-component--selections--subscribe form-group form-item form-item-submitted-selections-subscribe form-type-checkboxes form-group"><div id="edit-submitted-selections-subscribe" class="form-checkboxes"><div class="form-item form-item-submitted-selections-subscribe-subscribe form-type-checkbox checkbox"> <label class="control-label" for="edit-submitted-selections-subscribe-1"><input type="checkbox" id="edit-submitted-selections-subscribe-1" name="submitted[selections][subscribe][subscribe]" value="subscribe" class="form-checkbox" />订阅仲量联行物业资讯</label>
                    </div></div> <label class="control-label element-invisible" for="edit-submitted-selections-subscribe">订阅仲量联行物业资讯</label>
                    </div><div class="col-sm-8 Form__checkbox form-item webform-component webform-component-checkboxes webform-component--selections--terms-and-privacy form-group form-item form-item-submitted-selections-terms-and-privacy form-type-checkboxes form-group"><div id="edit-submitted-selections-terms-and-privacy" class="form-checkboxes"><div class="form-item form-item-submitted-selections-terms-and-privacy-termsandprivacy form-type-checkbox checkbox"> <label class="control-label" for="edit-submitted-selections-terms-and-privacy-1"><input required="required" type="checkbox" id="edit-submitted-selections-terms-and-privacy-1" name="submitted[selections][terms_and_privacy][termsAndPrivacy]" value="termsAndPrivacy" class="form-checkbox" />本人已阅读并同意仲量联行的 <a title="Terms of Use" class="terms_conditions_link">使用条款</a> 及 <a title="Privacy Policy" class="privacy_policy_link">隐私政策</a></label>
                    </div></div> <label class="control-label element-invisible" for="edit-submitted-selections-terms-and-privacy">Terms And Privacy <span class="form-required" title="This field is required.">*</span></label>
                    </div>  </div>
                      </fieldset>
                    <div class="form-item webform-component webform-component-hidden webform-component--elqcampaignid form-item form-item-submitted-elqcampaignid form-type-hidden" style="display: none"><input type="hidden" name="submitted[elqcampaignid]" value="" />
                    </div><div class="form-item webform-component webform-component-hidden webform-component--floorsize form-item form-item-submitted-floorsize form-type-hidden" style="display: none"><input type="hidden" name="submitted[floorsize]" value="" />
                    </div><div class="form-item webform-component webform-component-hidden webform-component--sourcelistingid form-item form-item-submitted-sourcelistingid form-type-hidden" style="display: none"><input type="hidden" name="submitted[sourcelistingid]" value="" />
                    </div><div class="form-item webform-component webform-component-hidden webform-component--originatingsite form-item form-item-submitted-originatingsite form-type-hidden" style="display: none"><input type="hidden" name="submitted[originatingsite]" value="dichandadang.com	" />
                    </div><div class="form-item webform-component webform-component-hidden webform-component--sourcespaceid form-item form-item-submitted-sourcespaceid form-type-hidden" style="display: none"><input type="hidden" name="submitted[sourcespaceid]" value="" />
                    </div><div class="form-item webform-component webform-component-hidden webform-component--olmtrackingid form-item form-item-submitted-olmtrackingid form-type-hidden" style="display: none"><input type="hidden" name="submitted[olmtrackingid]" value="" />
                    </div><div class="form-item webform-component webform-component-hidden webform-component--locale form-item form-item-submitted-locale form-type-hidden" style="display: none"><input type="hidden" name="submitted[locale]" value="zh-hans" />
                    </div><div class="form-item webform-component webform-component-hidden webform-component--enquirytype form-item form-item-submitted-enquirytype form-type-hidden" style="display: none"><input type="hidden" name="submitted[enquirytype]" value="Landlord form" />
                    </div><div class="form-item webform-component webform-component-hidden webform-component--address form-item form-item-submitted-address form-type-hidden" style="display: none"><input type="hidden" name="submitted[address]" value="" />
                    </div><div class="form-item webform-component webform-component-hidden webform-component--humandonotseethis form-item form-item-submitted-humandonotseethis form-type-hidden" style="display: none"><input type="hidden" name="submitted[humandonotseethis]" value="" />
                    </div><div class="form-item webform-component webform-component-hidden webform-component--keyword form-item form-item-submitted-keyword form-type-hidden" style="display: none"><input type="hidden" name="submitted[keyword]" value="" />
                    </div><div class="form-item webform-component webform-component-hidden webform-component--propertytype form-item form-item-submitted-propertytype form-type-hidden" style="display: none"><input type="hidden" name="submitted[propertytype]" value="" />
                    </div><div class="form-item webform-component webform-component-hidden webform-component--recipientemail form-item form-item-submitted-recipientemail form-type-hidden" style="display: none"><input type="hidden" name="submitted[recipientemail]" value="" />
                    </div><input type="hidden" name="details[sid]" />
                    <input type="hidden" name="details[page_num]" value="1" />
                    <input type="hidden" name="details[page_count]" value="1" />
                    <input type="hidden" name="details[finished]" value="0" />
                    <input type="hidden" name="form_build_id" value="form-EWT4uqKwvjomh_zweE4OsDBnJElo3etOvKuhG3mm49o" />
                    <input type="hidden" name="form_id" value="webform_client_form_212" />
                    <input type="hidden" name="elqFormName" value="cnConsolidatedPropertyForm" />
                    <input type="hidden" name="elqSiteId" value="362000045" />
                    <input type="hidden" name="elqDefaultTargetURL" value="" />
                    <input type="hidden" name="elqPost" value="" />
                    <input id="edit-elqCustomerGUID" type="hidden" name="elqCustomerGUID" value="" />
                    <input type="hidden" name="elqCookieWrite" value="0" />
                    <input type="hidden" name="user_headers" value="a:0:{}" />
                    <div class="form-actions"><button class="olm-btn btn-danger olm-btn-small bbtn form-submit" type="submit" name="op" value="提交">提交</button>
                    </div></form></div>    </article>
                    
                    </div>
                      </div>
                    </div>
                    

    
                    
                                  

</div>
</div>
</div>

<!--#include file="inc/footer.asp"-->