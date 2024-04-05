const fs = require("fs");
const {Document, Packer, Paragraph, TextRun} = require ("docx")
const {saveAs} = require ("file-saver")

fs.readFile("FadderApplications.csv", "utf-8", (err, data) => {
  if (err) console.log(err);
  else {
    const rows = data.split(/\r\n/)
    
    var questions = rows[0].split(",")

    rows.shift()

    var answers = []

    for (const row of rows) {
        const result = row.split(",")
        var obj = {}
      for(const index in result){
        const final = result[index].replace("{.}", ",")
        obj["q"+index] = final
      }
      answers.push(obj)
    }


  const creator = new documentCreator()
  const doc = creator.create([questions, answers])

    saveDocumentToFile(doc, 'testDocument.docx')
  }
})

class documentCreator {
  create([questions, answers]) {
    const document = new Document({
      sections: [{
          children: [
            ...answers.map((answer) => {
              const arr = []
              arr.push(this.createHeader(answer.q9, answer.q7, answer.q0))
                arr.push(this.createQA(questions[1], answer.q1, (answer.q1 != "No")))
                arr.push(this.createQA(questions[2], answer.q2, (answer.q2 != "First")))
                arr.push(this.createNewLine())
                if(answer.q3 != "") {
                  arr.push(this.createQA(questions[3], answer.q3))
                  arr.push(this.createNewLine())
                }
                if(answer.q4 != "") {
                  arr.push(this.createQA(questions[4], answer.q4))
                  arr.push(this.createNewLine())
                }
                arr.push(this.createQA(questions[5], answer.q5))
                arr.push(this.createNewLine())
                arr.push(this.createQA(questions[6], answer.q6))
                arr.push(this.createNewLine())
                arr.push(this.createQA(questions[8], answer.q8))
                arr.push(this.createNewLine())
                arr.push(this.createNewLine())
              return arr
            }).reduce((prev, curr) => prev.concat(curr), [])
          ],
      }],
  })
  
    return document
  }

  createHeader(name, klass, position) {
    return new Paragraph({
      children: [
        new TextRun({
          text: `${name} ${klass}, ${position}`,
          bold: true,
        })
      ]
    })
  }

  createNewLine() {
    return new Paragraph({
      children: [
        new TextRun("")
      ],
    })
  }

  createQA(question, answer, warning) {
    return new Paragraph({
      children: [
          new TextRun({
              text: question,
              bold: true
          }),
          new TextRun({
            text: answer,
            color: warning ? "FF0000" : "000000",
            break: 1
        })
      ],
    })
  }

}

function saveDocumentToFile(doc, fileName) {

  Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync(fileName, buffer)
  })

}