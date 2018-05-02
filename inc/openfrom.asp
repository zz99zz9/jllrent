<!--弹窗-->
<div class="intpop"><div class="close b_h">X</div>    <form>
      <span class="wrtit">请填写表格</span>
      <input type="text" class="wrint" value="您的姓名" id="wrint1"/><span class="errinfo"></span>
      <input type="text" class="wrint" value="手机号码" id="wrint2" onkeyup="this.value=this.value.replace(/\D/g,'')"  onafterpaste="this.value=this.value.replace(/\D/g,'')" max="11"/><span class="errinfo"></span>
      <input type="text" class="wrint" value="邮箱地址" id="wrint3"/><span class="errinfo"></span>
      <input type="button" class="wrbtn b_h" value="提交" id="wrbtn1"/>
      <input type="hidden" id="downurl" />
    </form></div>
<div class="mask"></div>