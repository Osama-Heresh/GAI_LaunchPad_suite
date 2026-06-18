import fs from 'node:fs';
import https from 'node:https';
import path from 'node:path';
import AdmZip from 'adm-zip';

const BRANCHES = ['main', 'master'];
const REPO = 'Osama-Heresh/Business_LaunchPad_Suite';
const TEMP_ZIP = path.join(process.cwd(), 'repo.zip');
const TEMP_DIR = path.join(process.cwd(), 'temp_import');

async function downloadFile(url: string, dest: string): Promise<boolean> {
  return new Promise((resolve) => {
    console.log(`Downloading: ${url}`);
    const file = fs.createWriteStream(dest);
    https.get(url, (response) => {
      if (response.statusCode === 301 || response.statusCode === 302) {
        // Handle redirect
        const redirectUrl = response.headers.location;
        if (redirectUrl) {
          downloadFile(redirectUrl, dest).then(resolve);
          return;
        }
      }

      if (response.statusCode !== 200) {
        console.error(`Failed to download: Status Code ${response.statusCode}`);
        file.close();
        fs.unlink(dest, () => {});
        resolve(false);
        return;
      }

      response.pipe(file);

      file.on('finish', () => {
        file.close();
        resolve(true);
      });
    }).on('error', (err) => {
      console.error(`Downloaded encountered error: ${err.message}`);
      fs.unlink(dest, () => {});
      resolve(false);
    });
  });
}

async function main() {
  let success = false;
  for (const branch of BRANCHES) {
    const url = `https://github.com/${REPO}/archive/refs/heads/${branch}.zip`;
    success = await downloadFile(url, TEMP_ZIP);
    if (success) {
      console.log(`Successfully downloaded branch ${branch}`);
      break;
    }
  }

  if (!success) {
    console.error('Failed to download repository from any known branches.');
    process.exit(1);
  }

  console.log(`Extracting zip to ${TEMP_DIR}...`);
  if (!fs.existsSync(TEMP_DIR)) {
    fs.mkdirSync(TEMP_DIR, { recursive: true });
  }

  try {
    const zip = new AdmZip(TEMP_ZIP);
    zip.extractAllTo(TEMP_DIR, true);
    console.log('Extraction complete!');
    fs.unlinkSync(TEMP_ZIP);
    
    // List contents of TEMP_DIR
    const topFiles = fs.readdirSync(TEMP_DIR);
    console.log('Extracted top-level contents:', topFiles);
  } catch (error) {
    console.error('Error during extraction:', error);
    process.exit(1);
  }
}

main();
