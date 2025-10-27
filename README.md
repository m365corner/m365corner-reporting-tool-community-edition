<h2>M365Corner Reporting Tool</h2>
<p>The M365Corner Reporting Tool is a Microsoft 365 reporting solution designed to simplify tenant management for administrators. Currently, it offers <b>25 customizable reports</b> that provide actionable insights into <b>Users, Groups, and Teams resources</b>. Reports can be exported and shared with other admins or stakeholders via email.</p>
<strong>Example Reports Include:</strong><br/>
<ul>
    <li>Disabled Users Report</li>
    <li>Unlicensed Users Report</li>
    <li>Empty Groups Report</li>
    <li>Archived Teams Report</li>
    <li>Team Owners Report</li>
</ul>
  
<strong>Prerequisites:</strong>

<strong>Creating the Entra ID or Azure AD App</strong>
<ul>
    <li>For more info: <a href="https://m365corner.com/m365-free-tools/create-azure-ad-app.html">https://m365corner.com/m365-free-tools/create-azure-ad-app.html</a></li>
</ul>

<strong>Registering App Details with M365Corner Reporting Tool</strong>
<ul>
    <li>
        Once you have the Entra ID or Azure AD app details [<strong>tenantId, clientId and clientSecret</strong>], you should add them to .env file, along with your <b>email credentials</b> [username and password <or security passcode, depending upon your email account security requirements>
            ], to register your tenant with M365Corner Reporting Tool.
    </li>
</ul>
<strong>Command for Running M365Corner Reporting Tool (within the project folder):</strong>
<ul>
    <li>node server.js</li>
</ul>

<strong>Packaging Command for M365Corner Reporting Tool (to run it as stand-alone application or exe)</strong>
<ul>
    <li>For Windows: pkg .  --output dist/reporting-tool --targets node18-win-x64</li>
    <li>For Linux: pkg . --output dist/reporting-tool --targets node18-linux-x64</li>
</ul>

<strong>Exe File Requirements</strong>
<p>The exe needs following files to function. These files and folders are to be copied into the /dist folder containing the generated exe file.</p>
<ul>
    <li><b>public folder</b> (or assets)</li>
    <li><b>.env file</b></li>
    <li><b>mocha.db file</b> (optional) – mocha.db gets generated when you run the exe. So, including mocha.db file in the /dist folder is optional.</li>
</ul>
    
   
<p><strong>Note:</strong> ensure you <b>turn off the anti-virus real-time scanning</b> before running this command. Otherwise, the exe gets quarantined, since it’s not code-signed project yet. You should also ensure the <b>anti-virus real-time scanning is turned off</b> every time you run this exe.</p>
<p>For more info on how to operate the reporting tool, read the <b>help documentation</b>: https://m365corner.com/m365-free-tools/community-edition-techdoc.docx</p>   
<p>For more info on <b>M365Corner Reporting Tool and the reports it has to offer</b>, along with their demos read: https://m365corner.com/m365-free-tools/m365-corner-reporting-tool.html</p>

