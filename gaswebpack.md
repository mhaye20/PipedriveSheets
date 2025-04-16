Creating the Bundle – using WebPack and Babel
Our next step is to use ES6 in Apps Script. We will use babel to cross compile from ES6 and will use webpack to create a bundle from the generated code.

I have a sample Apps Script Project here:

https://github.com/gsmart-in/AppsCurryStep1

Let’s look into the project structure.



The “server” sub-folder contains the code. api.js file contains the functions that we expose to Apps Script.

In the file lib.js we have es6 code. From this lib module, we can import other es6 files and npm packages.



We use webpack to bundle the code and babel to do the cross compilation.

Let us now look at the webpack.gas.js file:

This is the webpack configuration file. In summary, what this configuration file tells webpack is

 Compile the server/lib.js file in to Javascript compatible with old Javascript using babel. Then place the bundle in a folder “dist”
 Copy the file api.js without any change to output folder ‘dist’
 Copy some configuration files (appsscript.js and .clasp.json to the output folder ‘dist’)
One important thing to notice is these lines:

module.exports = {
  mode: 'development',
  entry:{
      lib:'./server/lib.js'
  },
  output: 
  {
    filename: '[name].bundle.js',
    path: path.resolve(__dirname, 'dist'),
    libraryTarget: 'var',
    library: 'AppLib'
  }
}
This means that webpack will expose a global variable AppLib through which you can access the classes and functions exported in the bundle.

Now see the file api.js.

function doGet() 
{
	var output = AppLib.getObjectValues();
	return ContentService.createTextOutput(output);
}
See server/lib.js file

function getObjectValues()
{
	let options = Object.assign({}, {source_url:null, header_row:1}, {content:"Hello, World"});

	return(JSON.stringify(options));
}

export {
    getObjectValues
}; 
We are using Object.assign() which is not supported by Apps Script. When babel cross-compiles lib.js, it will generate compatible code that works on Apps Script.

Let us now see package.json

{
  "name": "AppsPackExample1",
  "version": "1.0.0",
  "description": "",
  "main": "index.js",
  "scripts": {
    "gas": "webpack --config webpack.gas.js ",
    "deploy": "npm run gas && cd dist && clasp push && clasp open --webapp"
  },
  "keywords": [],
  "author": "",
  "license": "MIT",
  "devDependencies": {
    "@babel/core": "^7.4.0",
    "@babel/preset-env": "^7.4.2",
    "babel-loader": "^8.0.5",
    "copy-webpack-plugin": "^5.0.1",
    "webpack": "^4.29.6",
    "webpack-cli": "^3.3.0"
  },
  "dependencies": {
    "@babel/polyfill": "^7.4.0"
  }
}
When you run

$> npm run gas 
Webpack compiles and bundles the lib.js code (and any additional modules you have imported) in to a single javascript file and places it in the “dist” folder.

Then we can just use “clasp” to upload the code .

See the script “deploy” in package.json.

It runs webpack , then does “clasp push” and “clasp open”


Integrating npm modules with your Apps Script Project
One of the limiting features of Apps Script is that there is no easy way to integrate npm like packages to your project.

For example, you may want to use momentjs to play with date, or lodash utility functions in your script.

There, indeed, is a Library feature in Apps Script but that option has several limitations. We will not explore the library option of Apps Script in this article; we will install npm modules and bundle those modules using webpack to create an Apps Script compatible package.

Since we already started using webpack to create bundles that we can integrate to apps script, it should be easier now to add some npm packages. Let us start with momentjs

Open terminal, go to the AppsCurryStep1 folder you created in the last step and add momentjs to the mix

npm install moment --save
Now let us use some momentjs features in our Apps Script project.

Let us add a new function in lib.js

import * as moment from 'moment';
function getObjectValues()
{
	let options = Object.assign({}, {source_url:null, header_row:1}, {content:"Hello, World"});

	return(JSON.stringify(options));
}

function getTodaysDateLongForm()
{
	return moment().format('LLLL');
}

export {
    getObjectValues,
    getTodaysDateLongForm
};
Hint: Don’t forget to export the new function

Now let’s use this new function in api.js

function doGet() 
{
	var output = 'Today is '+AppLib.getTodaysDateLongForm()+"\n\n";
	
	return ContentService.createTextOutput(output);

}
Go to the command line and enter

npm run deploy
The updated script should open in the browser and print today’s date

There is not much fun getting todays date. Let us add another function that has a little more to do

function getDaysToAnotherDate(y,m,d)
{
	return moment().to([y,m,d]);
}
Now in api.js update doGet() and call getDaysToAnotherDate()

function doGet() 
{
	var output = 'Today is '+AppLib.getTodaysDateLongForm()+"\n\n";
	output += "My launch date is "+AppLib.getDaysToAnotherDate(2020,3,1)+"\n\n";

	return ContentService.createTextOutput(output);
}
Next, let us add lodash to the mix

First , run

  npm install lodash --save
Let us add a random number generator with the help of lodash

function printSomeNumbers()
{
	let out = _.times(6, ()=>
	{
		return _.padStart(_.random(1,100).toString(), 10, '.')+"\n\n"; 
	});

	return out;
}
Let us call this new function from api.js

function doGet() 
{
	var output = 'Today is '+AppLib.getTodaysDateLongForm()+"\n\n";
	output += "My launch date is "+AppLib.getDaysToAnotherDate(2020,3,1)+"\n\n";
	output += "\n\n";
	output += "Random Numbers using lodash\n\n";
	output += AppLib.printSomeNumbers();
	return ContentService.createTextOutput(output);

}
Deploy the project again

  npm run deploy