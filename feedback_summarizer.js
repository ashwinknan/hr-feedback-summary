const axios = require('axios');
const fs = require('fs');
const ExcelJS = require('exceljs');
require('dotenv').config();

// Load the feedback data
const workbook = new ExcelJS.Workbook();
workbook.xlsx.readFile('feedback.xlsx')
  .then(function() {
    const worksheet = workbook.getWorksheet(1);
    let feedbackData = {};

    worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
      if (rowNumber !== 1) {
        const personReceivingFeedback = row.getCell(2).value;
        const feedbackDescription = row.getCell(3).value;
        if (feedbackData[personReceivingFeedback]) {
          feedbackData[personReceivingFeedback].push(feedbackDescription);
        } else {
          feedbackData[personReceivingFeedback] = [feedbackDescription];
        }
      }
    });

    // Process the feedback for each person
    let summarizedFeedback = {};
    for (let person in feedbackData) {
      const text = feedbackData[person].join("\n");
      const strengthsPrompt = `I have the following feedback for ${person}:\n\n${text}\n\nGiven these points, please provide a comprehensive yet concise summary of the person's strengths. These should include qualities or behaviors that positively contribute to the organization's business goals and create a good culture. Provide a maximum of three points, numbered 1, 2, and 3 respectively. Prioritize the most important and impactful strengths.`;
      const improvementsPrompt = `I have the following feedback for ${person}:\n\n${text}\n\nGiven these points, please provide a comprehensive yet concise summary of the areas of improvement for the person. These should include qualities or behaviors that are detrimental to the organization's business goals or culture, setting a bad example for others. Provide a maximum of three points, numbered 1, 2, and 3 respectively. Prioritize the most important and critical areas of improvement.`;

      // Send a request to the OpenAI API
      axios.post('https://api.openai.com/v1/completions', {
        model: "text-davinci-002",
        prompt: strengthsPrompt,
        max_tokens: 60,
      }, {
        headers: {
          'Authorization': `Bearer ${process.env.OPENAI_API_KEY}`,
          'Content-Type': 'application/json'
        }
      }).then((response) => {
        const strengthsSummary = response.data.choices[0].text.trim();

        axios.post('https://api.openai.com/v1/completions', {
          model: "text-davinci-002",
          prompt: improvementsPrompt,
          max_tokens: 60,
        }, {
          headers: {
            'Authorization': `Bearer ${process.env.OPENAI_API_KEY}`,
            'Content-Type': 'application/json'
          }
        }).then((response) => {
          const improvementsSummary = response.data.choices[0].text.trim();
          summarizedFeedback[person] = {
            strengths: strengthsSummary,
            improvements: improvementsSummary
          };

          // Once all feedback is processed, write the summaries to an Excel file
          if (Object.keys(summarizedFeedback).length === Object.keys(feedbackData).length) {
            const outputWorkbook = new ExcelJS.Workbook();
            const outputWorksheet = outputWorkbook.addWorksheet('Summarized Feedback');
            outputWorksheet.columns = [
              { header: 'Person', key: 'person' },
              { header: 'Summary of Strengths', key: 'strengths' },
              { header: 'Summary of Areas of Improvements', key: 'improvements' }
            ];
            for (let person in summarizedFeedback) {
              outputWorksheet.addRow({
                person: person,
                strengths: summarizedFeedback[person].strengths,
                improvements: summarizedFeedback[person].improvements
              });
            }
            outputWorkbook.xlsx.writeFile('output.xlsx');
          }
        });
      });
    }
  });
