# Office-365-Outlook-Add-in-with-Angular

**Create an Outlook Add-in Project** 
`npm install -g yo generator-office` \
`yo office`

* _Choose a project type:_
    Office Add-in Task Pane project using Angular framework

 * _Choose a script type:_ 
     Typescript
    
 * _What do you want to name your add-in (My Office Add-in):_
   testAddin

 * _Which Office client application would you like to support:_ 
   Outlook


**Create an Angular Project**\
`npm install -g @angular/cli`\
`ng new testAngular`\
`cd testAngular`

**Transfer Files from testAddin to testAngular**
* (Optional) Copy all the images from the _assets_ folder of _testAddin_ to _src/assets_ folder of _testAngular_
* Copy the manifest.xml from “outlook-add-in” and paste it in the root folder of “demo-angular-addin”

**Make changes in Manifest file**
* Find _taskpane.html_ in the manifest file and replace it with _index.html_
* Find _https://localhost:3000_ in the manifest file and replace it with _https://localhost:4200_. The URL should start with _https_ not _http_.

**Install dependencies** \
`npm install --save @microsoft/office-js` \
`npm install --save @microsoft/office-js-helpers` \
`npm install --save office-ui-fabric-js`

**Install dev dependencies**
`npm install --save-dev @types/office-js`\
`npm install --save-dev @types/office-runtime`\
`npm install --save-dev office-addin-debugging`\
`npm install --save-dev office-addin-dev-certs`\
`npm install --save-dev office-toolbox`

**Update the tsconfig.app.json**
* Add _office-js_ under the types array under _compilerOptions_ - _"types": [ "office-js"]_ 
* Add the new _target_ under _compilerOptions_ - _"target": "es5"_

![image](https://user-images.githubusercontent.com/47311938/219957915-cdf805bd-7ef2-4c95-92fe-3ee1469086f7.png)

**Update index.html**\
`<script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>` 

`<link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css" />` 

`<link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.components.min.css" />` 

`<script src="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/js/fabric.min.js"></script>`

![image](https://user-images.githubusercontent.com/47311938/219958085-c356920f-265b-4641-ae17-a757bdc2da24.png)

**Update main.ts** \
`Office.initialize = function () {
  platformBrowserDynamic().bootstrapModule(AppModule)
  .catch(err => console.error(err));
}`

![image](https://user-images.githubusercontent.com/47311938/219958178-d1b837fc-3252-4041-b9b8-9911674fcd2d.png)

**Update app-routing.module.ts**
* Add _{ useHash: true }_ in imports array
![image](https://user-images.githubusercontent.com/47311938/219958240-c5d1fd72-6ab4-458d-bdd5-f0a8052a4bf1.png)










