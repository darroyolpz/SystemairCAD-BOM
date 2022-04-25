# SystemairCAD BOM

## Background

Spare parts and service represent a big chunk of the revenue of a company. In fact, many companies sacrifies profits in the product sold in exchange for future concurrent revenue due to maintenance, spare parts, etc.

This business model can be a game-changer for many companies fighting for medium and big projects, but they also have to be efficient at quoting their products, and this is where the situation gets tricky for many of them.

Quoting is a time-consuming job when done manually, but we can speed-up the process using the data we already have in our system.

## SystemairCAD Data

In our case, we use SystemairCAD to calculate the air handling units. It's a very good software, quite visual and fast. It has an option to check the items in each unit that we have calculated, so it's more than enough for our purpose.

![SystemairCAD](https://raw.githubusercontent.com/darroyolpz/SystemairCAD-BOM/master/img/SystemairCAD_Export_BOM.jpg)

This is the data we get once that button is pressed. We have each unit in each spreadsheet, with all its item numbers and prices.

![BOM](https://raw.githubusercontent.com/darroyolpz/SystemairCAD-BOM/master/img/BOM.jpg)

## The Code

The code we need to write is very straight-forward. Copy all the data from each sheet and paste it into a new table, so that we can manipulate it later.

Also add the name of the AHU, so that we can filter it just in case.

## Benefits

Not only we can quickly quote spares, but also provide this information to production and purchase, to negotiate with suppliers and improve overall contribution margin for the product, and avoiding possible delays in manufacturing.