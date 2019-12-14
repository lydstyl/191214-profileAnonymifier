const fs = require('fs');
const path = require('path');
const scrapedin = require('scrapedin');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

function scrapeProfileAndMakeResume() {
  if (fs.existsSync('./profile.json')) {
    writeDocx();
  } else {
    scrapedin({
      email: process.env.EMAIL,
      password: process.env.PASSWORD
    })
      .then(profileScraper =>
        // profileScraper('https://www.linkedin.com/in/gabrielbrun/')
        profileScraper('https://www.linkedin.com/in/elodiebattaire/')
      )
      .then(profile => {
        fs.writeFile(
          'profile.json',
          JSON.stringify(profile, null, 2),
          writeDocx
        );
      });
  }
}

function writeDocx() {
  const profile = require('./profile.json');
  console.log(profile);

  //Load the docx file as a binary
  const content = fs.readFileSync(
    path.resolve(__dirname, 'cvTemplate.docx'),
    'binary'
  );

  const zip = new PizZip(content);

  const doc = new Docxtemplater();
  doc.loadZip(zip);

  //set the templateVariables
  doc.setData({
    name: profile.profileAlternative.name,
    headline: profile.profileAlternative.headline
  });

  try {
    // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
    doc.render();
  } catch (error) {
    // The error thrown here contains additional information when logged with JSON.stringify (it contains a properties object).
    var e = {
      message: error.message,
      name: error.name,
      stack: error.stack,
      properties: error.properties
    };
    console.log(JSON.stringify({ error: e }));
    if (error.properties && error.properties.errors instanceof Array) {
      const errorMessages = error.properties.errors
        .map(function(error) {
          return error.properties.explanation;
        })
        .join('\n');
      console.log('errorMessages', errorMessages);
      // errorMessages is a humanly readable message looking like this :
      // 'The tag beginning with "foobar" is unopened'
    }
    throw error;
  }

  const buf = doc.getZip().generate({ type: 'nodebuffer' });

  // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
  fs.writeFileSync(path.resolve(__dirname, 'cv.docx'), buf);
  console.log('finish');
}

scrapeProfileAndMakeResume();
