// Import Tinytest from the tinytest Meteor package.
import { Tinytest } from "meteor/tinytest";

// Import and rename a variable exported by microsoft.js.
import { name as packageName } from "meteor/uibalabs:microsoft";

// Write your tests here!
// Here is an example.
Tinytest.add('microsoft - example', function (test) {
  test.equal(packageName, "microsoft");
});
