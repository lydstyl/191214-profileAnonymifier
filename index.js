const fs = require('fs');
const path = require('path');
const dotenv = require('dotenv');
const scrapedin = require('scrapedin');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

const jsonProfile = `./profiles/${process.env.PROFIL}.json`;

function scrapeProfileAndMakeResume() {
  if (fs.existsSync(jsonProfile)) {
    writeDocx();
  } else {
    scrapedin({
      email: process.env.EMAIL,
      password: process.env.PASSWORD
    })
      .then(profileScraper =>
        profileScraper(`https://www.linkedin.com/in/${process.env.PROFIL}/`)
      )
      .then(profile => {
        fs.writeFile(jsonProfile, JSON.stringify(profile, null, 2), writeDocx);
      });
  }
}

function writeDocx() {
  const profile = require(jsonProfile);

  const content = fs.readFileSync(
    path.resolve(__dirname, 'cvTemplate.docx'),
    'binary'
  );
  const zip = new PizZip(content);
  const doc = new Docxtemplater();
  doc.loadZip(zip);

  // console.log(profile.educations);

  doc.setData({
    name: profile.profileAlternative.name || 'CV',
    headline: profile.profileAlternative.headline,
    description: profile.profile.summary,
    skills: profile.skills,
    experiences: profile.positions.map(xp => ({
      begin: xp.date1 ? xp.date1 : '',
      job: xp.title,
      company: xp.companyName
    })),
    educations: profile.educations
  });

  try {
    doc.render();
  } catch (error) {
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
    }
    throw error;
  }

  const buf = doc.getZip().generate({ type: 'nodebuffer' });

  fs.writeFileSync(
    path.resolve(__dirname, `./profiles/${process.env.PROFIL}.docx`),
    buf
  );
  console.log('finish');
}

dotenv.config();

scrapeProfileAndMakeResume();
