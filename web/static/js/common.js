function loginCheck(){
   
    const sessionInsName=document.getElementById('sessionInsName').innerText;
    const sessionInsId=document.getElementById('sessionInsId').innerText;
    const insIsLogin=localStorage.getItem('isLogin')
    console.log(window.location.href);
    if(insIsLogin!=='true'){
        window.location.href='/';
    }
    if(sessionInsName.includes('None')){
        localStorage.removeItem('isLogin');
        window.location.href='/';
    }

    return 'isLogin'
   
}

function logout(){
    localStorage.removeItem('isLogin');
    window.location.href='/logout';
}

function goIndex(){
    window.location.href='/index';
}

function hideInsSelectBlockAndGetInsInfo(id){
    
    const sessionInsId=document.getElementById('sessionInsId').textContent;
    const sessionInsName=document.getElementById('sessionInsName').textContent;
    const sessionRoleBlock=document.getElementById('sessionInsRole');
    const showRole=document.getElementById('showRole');

    const sessionInsRole=sessionRoleBlock.textContent;
    if(id){
        const insSelectBlock=document.getElementById(id);
        if(sessionInsRole==='admin'){
            insSelectBlock.style.display='block';
            showRole.innerText='管理员'
        }else if(sessionInsRole==='ins'){
            insSelectBlock.style.display='none';
            showRole.innerText='教练'
        }
    }else{
        if(sessionInsRole==='admin'){   
            showRole.innerText='管理员'
        }else if(sessionInsRole==='ins'){ 
            showRole.innerText='教练'
        }
    }
    sessionRoleBlock.style.display='none';
    return ({'sessionInsId':sessionInsId,'sessionInsName':sessionInsName,'sessionInsRole':sessionInsRole})
}


function generateInsList(insData,selectId){
    const insSelect=document.getElementById(selectId);
    insData.forEach(ins=>{
        const option=document.createElement('option')
        option.value=ins.slice(0,8);
        option.text=ins.slice(8,);
        insSelect.appendChild(option);                    
        });
    }

function selectToday(id,format){
    const today = new Date();
    let formattedDate;
    // 将日期格式化为 yyyy-mm-dd 的形式
    if(format==='dateTime'){
        const year = today.getFullYear();
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const day = String(today.getDate()).padStart(2, '0');
        const hours = String(today.getHours()).padStart(2, '0');
        const minutes = String(today.getMinutes()).padStart(2, '0');
        formattedDate = `${year}-${month}-${day}T${hours}:${minutes}`;
    }else if(format==='date'){
        const year = today.getFullYear();
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const day = String(today.getDate()).padStart(2, '0');
        // const hours = String(today.getHours()).padStart(2, '0');
        // const minutes = String(today.getMinutes()).padStart(2, '0');
        formattedDate = `${year}-${month}-${day}`
    }

    // 将日期设置为输入框的默认值
    document.getElementById(id).value = formattedDate;
}

function dateToString(dateInput,format){
    const today = new Date(dateInput);
    let formattedDate;
    // 将日期格式化为 yyyy-mm-dd 的形式
    if(format==='dateTime'){
        const year = today.getFullYear();
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const day = String(today.getDate()).padStart(2, '0');
        const hours = String(today.getHours()).padStart(2, '0');
        const minutes = String(today.getMinutes()).padStart(2, '0');
        formattedDate = `${year}-${month}-${day}T${hours}:${minutes}`;
    }else if(format==='date'){
        const year = today.getFullYear();
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const day = String(today.getDate()).padStart(2, '0');
        // const hours = String(today.getHours()).padStart(2, '0');
        // const minutes = String(today.getMinutes()).padStart(2, '0');
        formattedDate = `${year}-${month}-${day}`
    }else if(format==='time'){
        const hours = String(today.getHours()).padStart(2, '0');
        const minutes = String(today.getMinutes()).padStart(2, '0');
        formattedDate = `${hours}:${minutes}`;
    }
    return formattedDate;
}

function calculateDate(dateInput,days){
    const currentDate = new Date(dateInput);

    // 获取当前日期中的日期部分
    const day = currentDate.getDate();

    // 将日期部分设置为当前日期加上30天的日期
    currentDate.setDate(day + days);

    // 将日期转换为字符串并输出
    const formattedDate = currentDate.toLocaleDateString();
    
    return formattedDate;
}

function selectDate(dateInput,id,format){
  const today = new Date(dateInput);
  let formattedDate;
  // 将日期格式化为 yyyy-mm-dd 的形式
  if(format==='dateTime'){
      const year = today.getFullYear();
      const month = String(today.getMonth() + 1).padStart(2, '0');
      const day = String(today.getDate()).padStart(2, '0');
      const hours = String(today.getHours()).padStart(2, '0');
      const minutes = String(today.getMinutes()).padStart(2, '0');
      formattedDate = `${year}-${month}-${day}T${hours}:${minutes}`;
  }else if(format==='date'){
      const year = today.getFullYear();
      const month = String(today.getMonth() + 1).padStart(2, '0');
      const day = String(today.getDate()).padStart(2, '0');
      // const hours = String(today.getHours()).padStart(2, '0');
      // const minutes = String(today.getMinutes()).padStart(2, '0');
      formattedDate = `${year}-${month}-${day}`
  }

  // 将日期设置为输入框的默认值
  document.getElementById(id).value = formattedDate;
}

function dateFormat(currentDate,fmt){
    let formattedDate;
    if(fmt==='date'){
        var year = currentDate.getFullYear();
        var month = ("0" + (currentDate.getMonth() + 1)).slice(-2);
        var day = ("0" + currentDate.getDate()).slice(-2);
        formattedDate = year + "-" + month + "-" + day;
    }else if(fmt==='time'){
        const hours = String(today.getHours()).padStart(2, '0');
        const minutes = String(today.getMinutes()).padStart(2, '0');
        const seconds = currentDate.getSeconds();
        // const formattedDatetime = `${year}-${month}-${day}T${hours}:${minutes}`;
        formattedDate = `${hours}:${minutes}:${seconds}`;
    }
    
    // 拼接日期字符串，例如：YYYY-MM-
    if (formattedDate.includes('NaN')){
        return '-';
    }else{
        return formattedDate;
    }              
}

  // 显示自定义模态框
  function showCustomModal() {
    const modal = document.getElementById('customModal');
    modal.style.display = 'block';
  }

  // 隐藏自定义模态框
  function hideCustomModal() {
    const modal = document.getElementById('customModal');
    modal.style.display = 'none';
  }

  // 确认或取消操作
  function confirmAction(isConfirmed) {
    if (isConfirmed) {
      // 执行确认操作
      console.log('确认');
    } else {
      // 执行取消操作
      console.log('取消');
    }

    // 隐藏模态框
    hideCustomModal();
  }


  class customButton{
    constructor(id='customModel'){
        this.className='customButton';
        this.id=id;
    }

      // 显示自定义模态框
    showCustomModal() {
        const modal = document.getElementById(this.id);
        modal.style.display = 'block';
    }

    // 隐藏自定义模态框
    hideCustomModal() {
        const modal = document.getElementById(this.id);
        modal.style.display = 'none';
    }

    // 确认或取消操作
    confirmAction(isConfirmed) {
        if (isConfirmed) {
        // 执行确认操作
        console.log('确认');
        } else {
        // 执行取消操作
        console.log('取消');
        }

        // 隐藏模态框
        hideCustomModal(this.id);
    }
  }