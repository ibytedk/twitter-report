function generateHashtagReport() {
  // Access active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Access the sheet containing the posts and impressions
  // Assuming it's the first one
  var sheet = ss.getSheets()[0];
  
  // Get data from the sheet
  var data = sheet.getDataRange().getValues();
  
  // Object to hold hashtags and their corresponding impressions
  var hashtags = {};
  
  // Loop over the data
  data.forEach(function(row) {
    // Assuming posts are in first column and impressions in second
    var post = row[2];
    console.log(post);
    var impressions = row[4];
    
    // If post is not a string, convert it to string
    if (typeof post !== 'string') {
      post = String(post);
    }
    
    // Extract hashtags from the post
    var matches = post.match(/#[a-zA-Z0-9_]+/g);
    
    // If there are hashtags, add their impressions to the object
    if(matches){
      matches.forEach(function(hashtag){
        if(hashtags[hashtag]){
          hashtags[hashtag].impressions += impressions;
          hashtags[hashtag].count += 1;
        } else {
          hashtags[hashtag] = {
            impressions: impressions,
            count: 1
          }
        }
      });
    }
  });
  
  // Create a new sheet for the report
  var report = ss.insertSheet('Hashtag Report');
  
  // Prepare data for the report
  var reportData = Object.keys(hashtags).map(function(hashtag) {
    return [hashtag, hashtags[hashtag].impressions, hashtags[hashtag].count, hashtags[hashtag].impressions / hashtags[hashtag].count ];
  });

  // Filter the reportData to include only tags with count >= 50
  reportData = reportData.filter(function(tag) {
    return tag[2] >= 25;
  });
  
  // Sort data by impressions
  reportData.sort(function(a, b) {
    return b[1] - a[1];
  });
  
  // Add headers to the report data
  reportData.unshift(['Hashtag', 'Total Impressions', 'Count', 'Average']);
  
  // Populate the report sheet with the report data
  report.getRange(1, 1, reportData.length, reportData[0].length).setValues(reportData);

  // Create a filter on the data range
  report.getDataRange().createFilter();
}

