var express = require('express');
var router = express.Router();
var authHelper = require('../helpers/auth');
var graph = require('@microsoft/microsoft-graph-client');

const excelToJson = require('convert-excel-to-json');






/* GET /calendar */
router.get('/', async function(req, res, next) {
  let parms = { title: 'Excel' };

  const accessToken = await authHelper.getAccessToken(req.cookies, res);
  const userName = req.cookies.graph_user_name;

  if (accessToken && userName) {
    parms.user = userName;

    // Initialize Graph client
    const client = graph.Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      }
    });

    // Set start of the calendar view to today at midnight
    const start = new Date(new Date().setHours(0,0,0));
    // Set end of the calendar view to 7 days from start
    const end = new Date(new Date(start).setDate(start.getDate() + 7));
    
    try {
      // Get the first 10 events for the coming week
      console.log("hi");
      const result = await client
      .api('/me/drive/sharedWithMe')
     // .api('/me/drive/items/01V3NFNIAXGLGAH7PDPVB2FTH5MHU7NQ22/workbook/worksheets/testsheet/usedRange(valuesOnly=true)')
      //.top(10)
      //.select('id,name,webUrl')
      //.orderby('start/dateTime DESC')
      .get();

       console.log(result.value[1]);
      // parms.events = result.value;
      // res.render('excel', parms);


      const resultexcel = excelToJson({
        sourceFile: './public/QnA Employee Handbook - Dev Version.xlsx'
      });
      
      

      parms.excel = resultexcel.Sheet1;
      console.log(parms);
      res.render('excel', parms);

      

    } catch (err) {
      parms.message = 'Error retrieving events';
      parms.error = { status: `${err.code}: ${err.message}` };
      parms.debug = JSON.stringify(err.body, null, 2);
      res.render('error', parms);
    }

  } else {
    // Redirect to home
    res.redirect('/');
  }
});

module.exports = router;