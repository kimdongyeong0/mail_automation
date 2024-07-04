function createAndSendForms() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  
  var courseFormLinks = {};
  var courseEmails = {};

  var processedData = []; // 처리된 데이터 배열
  var skippedData = []; // 누락된 데이터 배열
  var sentEmails = []; // 전송된 메일 배열

  // 헤더 행 건너뛰기
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var studentName = row[0];  // A열
    var email = row[1];        // B열
    var courseName = row[2];   // C열
    
    if (!courseName) {
      // 누락된 데이터를 skippedData 배열에 추가
      skippedData.push({
        rowNumber: i+1,
        studentName: studentName,
        email: email,
        reason: "Missing course name" // 코스 이름이 누락된 이유
      });
      continue;
    }

    // 처리된 데이터를 배열에 추가
    processedData.push({
      rowNumber: i+1,
      studentName: studentName,
      email: email,
      courseName: courseName
    });
    
    // 각 코스별 이메일 수집
    if (!courseEmails[courseName]) {
      courseEmails[courseName] = [];
    }
    courseEmails[courseName].push(email);
    
    // 해당 코스에 대한 폼이 없으면 폼 생성
    if (!courseFormLinks[courseName]) {
      var form = FormApp.create(courseName + " Course Evaluation");
      form.setDescription("Please complete the evaluation for " + courseName + ".");
      
      // 이메일 주소를 수집하지 않도록 설정하고 누구나 응답할 수 있도록 허용
      form.setCollectEmail(false);
      form.setRequireLogin(false);
      
      // 몇 가지 예제 질문 추가 (이 질문들을 사용자 정의할 수 있습니다)
      form.addScaleItem()
        .setTitle('How would you rate this course overall?')
        .setBounds(1, 5)
        .setLabels('Poor', 'Excellent')
        .setRequired(true);
      
      form.addParagraphTextItem()
        .setTitle('What aspects of the course did you find most valuable?');
      
      form.addParagraphTextItem()
        .setTitle('Do you have any suggestions for improving the course?');

      form.addMultipleChoiceItem()
        .setTitle('Which of the following best describes your experience with this course?')
        .setChoiceValues([
          'Exceeded my expectations',
          'Met my expectations',
          'Somewhat below my expectations',
          'Far below my expectations'
        ])
        .setRequired(true);
      
      var formUrl = form.getPublishedUrl();
      courseFormLinks[courseName] = formUrl;
    }
  }

  // 처리된 데이터 로그 출력
  console.log("Processed Data:");
  console.log(JSON.stringify(processedData, null, 2));

  // 누락된 데이터 로그 출력
  console.log("\nSkipped Data:");
  console.log(JSON.stringify(skippedData, null, 2));

  // 코스별 이메일 수 로그 출력
  console.log("\nCourse Email Counts:");
  for (var courseName in courseEmails) {
    console.log(courseName + ": " + courseEmails[courseName].length + " students");
  }

  // 이메일 전송
  for (var courseName in courseEmails) {
    var emails = courseEmails[courseName].join(',');
    var formUrl = courseFormLinks[courseName];
    var mailSubject = courseName + ' Course Evaluation';
    var mailBody = 'Dear student,\n\n' +
                   'Please complete the evaluation for ' + courseName + ' using the following link:\n\n' + 
                   formUrl + '\n\n' +
                   'This survey is anonymous and does not require you to sign in.\n\n' +
                   'Your feedback is important to us and will help improve the course for future students.\n\n' +
                   'Thank you for your participation!\n\n' +
                   'Best regards,\n' +
                   'The Course Administration Team';
    // var htmlBody = 'Dear student,<br><br>' +
    //                'Please complete the evaluation for ' + courseName + ' using the following link:<br><br>' + 
    //                '<a href="' + formUrl + '">' + formUrl + '</a><br><br>' +
    //                'This survey is anonymous and does not require you to sign in.<br><br>' +
    //                '<span style="color:red; font-weight:bold;">Your feedback is important to us and will help improve the course for future students.</span><br><br>' +
    //                'Thank you for your participation!<br><br>' +
    //                'Best regards,<br>' +
    //                'The Course Administration Team';               
    
    try {
      MailApp.sendEmail({
        bcc: emails,
        subject: mailSubject,
        body: mailBody
        // htmlBody : htmlBody
      });

      // 전송된 이메일 데이터 저장
      sentEmails.push({
        courseName: courseName,
        recipientCount: courseEmails[courseName].length,
        subject: mailSubject,
        formUrl: formUrl,
        sentAt: new Date().toISOString()
      });

      console.log(courseName + ": Sent to " + courseEmails[courseName].length + " students");
    } catch (error) {
      console.error("Error sending email for " + courseName + ": " + error.message);
    }
  }
}
