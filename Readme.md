<h1 align="center">Excel Mail Merge To Email</h1>

## 💻 About
<p align="center">A tool to allow you to create mail merge emails from an Excel table on OneDrive. e.g. put a name in the body of each email.</p>

## ℹ️ How To Use It
<p align="center">
 <a href="https://trctest.azurewebsites.net">Go Here To Try It and Follow The Instructions</a>

</p>

## 🗂️ Tech:
<p align="left">
I wanted to demonstrate what can be done with Blazor WASM and some of the great tools available for it. The project outside of the API calls is a static site (no backend).
</p>
<ul>
  <li>MudBlazor component libary has been used for the UI which is great for blazor development</li>
  <li>JSInterop both ways is demonstrated with the use of the OneDrive file picker</li>
  <li>Graph has been used to process an excel table and use the data</li>
  <li>Graph has been used to send emails in bulk</li>
  <li>I intend to demontrate TinyMCE or CKEditor working in blazor wasm as a HTML editor for the email body</li>
</ul>

## 📥 Installation & Set Up
<p> Clone the repo and create an App Registration in Azure Active Directory. It will need the permissions User.Read,Mail.Send,Files.Read,Files.Read.All,Sites.Read.All. It will also need Access Tokens/ID Tokens selecting alogn with a relevant rspa redirect url.</p>
<p> Set your Client ID in appsettings.Json</p>
<hr>

## TODO
<p>I ran out of time as I only saw this 2 days ago</p>
<ul>
  <li>Improve the documentation here and instructions on the site</li>
  <li>Look into batching the emails to save on graph calls and throttling</li>
  <li>Expose exceptions where the selected Excel data isn't valid to the user</li>
  <li>General UI improvements such as progress of sending the emails</li>
  <li>Change body text field to a HTML editor such as TinyMCE</li>
</ul>
