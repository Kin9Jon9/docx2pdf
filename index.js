const libre = require('libreoffice-convert');
const path = require('path');
const fs = require('fs');
const extend = '.pdf'
const async = require('async');

//파일명에 경로를 합칩니다.
const getEnterPath = (filename) => {
  return path.join(__dirname, `/docx2/${filename}`);
}

//파일명에 경로를 합치되, 기존의 .docx 확장자는 삭제합니다.
const getOutputPath = (filename) => {
  filename = filename.substring(0,filename.length-5)
  return path.join(__dirname, `/pdf/${filename}${extend}`);
}

// Docx2Pdf 함수.
function processing(enterPath, outputPath){
  return new Promise(async (resolve, reject) => {
    const file = fs.readFileSync(enterPath);
    await libre.convert(file, extend, undefined, (err, done) => {
        if (err) {
          console.log(`변환 중 에러 : ${err}`);
        } else {
          fs.writeFileSync(outputPath, done);
          resolve(true);
        }
    });
  })
}

// docx 폴더 내부에 있는 모든 파일을 가져옵니다.
function getDocxList(){
  return new Promise((resolve, reject)=>{
    fs.readdir(path.join(__dirname, '/docx2'),(err, fileList)=>{
      if (err) reject(err);
      else resolve(fileList);
    })
  })
}

getDocxList().then( async (result) => {
    //파일을 하나씩 변환합니다.
    for(const index in result){
      enterPath = getEnterPath(result[index]);
      outputPath = getOutputPath(result[index]);
      await processing(enterPath, outputPath)
      console.log(`${result.length} 개의 파일 중 ${ Number(index)+1 }번 째 파일인 ${result[index]}을 처리 중 입니다.`)
    }
  }, (error) => {
    console.log(error);
  }
  );