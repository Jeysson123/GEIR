// Unit tests with Node's built-in runner (node:test): no extra dependencies, so they run
// the same on a developer PC, the VM and the CI pipeline. `npm run test:unit`.
import './task-lines.test';
import './requirements.test';
import './estimates.test';
import './email.test';
import './workbook.test';
import './sprint.test';
import './microsoft-token.test';
import './build.test';
import './upload.test';
