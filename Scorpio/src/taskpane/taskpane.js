Office.onReady(info => {
  if (info.host === Office.HostType.OneNote) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("btn-refresh").onclick = run;
    run();
  }
});

function download(filename, text) {
  var element = document.createElement('a');
  element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(text));
  element.setAttribute('download', filename);

  element.style.display = 'none';
  document.body.appendChild(element);

  element.click();

  document.body.removeChild(element);
}

export async function run() {
  try {
    await OneNote.run(async context => {

      // Get the current page.
      var page = context.application.getActivePage();
      var pageContents = page.contents;
      var contentDiv = document.getElementById("page-content")
      var outlines = [];
      var paragraphs = [];
      var html_objects = [];
      var html_content = "";
      // Queue a command to load the pageContents to access its data.
      context.load(pageContents);
      page.load("title");

      // Run the queued commands, and return a promise to indicate task completion.
      return context.sync()
        .then(function () {
          for (var i = 0; i < pageContents.items.length; i++) {
            var pageContent = pageContents.items[i];
            if (pageContent.type === "Outline") {
              pageContent.outline.load("id, paragraphs/items, paragraphs/type")
              outlines.push(pageContent.outline);
            }
          }
          // Add Page Title
          html_content += "<h1>" + page.title + "</h1>";
          return context.sync()
        })
        .then(function () {
          for (var i = 0; i < outlines.length; i++) {
            console.log("Outline ID: " + outlines[i].id);
              for (var j = 0; j < outlines[i].paragraphs.items.length; j++) {
                var paragraph = outlines[i].paragraphs.items[j];
                if (paragraph.type == "RichText") {
                  paragraphs.push(paragraph);
                  var html = paragraph.richText.getHtml();
                  html_objects.push(html);
                  paragraph.load("richtext/id, richtext/text");
                }
              }
          }
          return context.sync()
        })
        .then(function () {
          for (var i = 0; i < html_objects.length; i++) {
            //console.log("RichText ID: " + paragraphs[i].id);
            var html = html_objects[i].value;
            // Process code block
            var html_element = $($.parseHTML(html));
            if (html_element.css('font-family') === 'Consolas') {
              html = "<pre><code>" + html_element.text() + "</code></pre>";
            }
            html_content += html;
          }
          var turndownService = new TurndownService({ 
            headingStyle: 'atx',
            codeBlockStyle: 'fenced'
          })
          var markdown = turndownService.turndown(html_content);
          // Save the markdown to textarea
          $("#markdown-editor").val(markdown);
        });
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}
