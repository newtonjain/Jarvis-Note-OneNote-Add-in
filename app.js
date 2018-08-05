/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

'use strict';

(function () {

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#run').click(run);
            setInterval(run, 3000);
        });
    };

    function run() {
        return OneNote.run(function (context) {
            console.log('/////////////');
            // Get the collection of pageContent items from the page.
            var pageContents = context.application.getActivePage().contents;

            // Get the first PageContent on the page, and then get its outline's paragraphs.
            var outlinePageContents = [];
            var paragraphs = [];
            var richTextParagraphs = [];
            // Queue a command to load the id and type of each page content in the outline.
            pageContents.load("id,type");

            // Run the queued commands, and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    // Load all page contents of type Outline
                    $.each(pageContents.items, function (index, pageContent) {
                        if (pageContent.type == 'Outline') {
                            pageContent.load('outline,outline/paragraphs,outline/paragraphs/type');
                            outlinePageContents.push(pageContent);
                        }
                    });
                    return context.sync();
                })
                .then(function () {
                    // Load all rich text paragraphs across outlines
                    $.each(outlinePageContents, function (index, outlinePageContent) {
                        var outline = outlinePageContent.outline;
                        paragraphs = paragraphs.concat(outline.paragraphs.items);
                    });
                    $.each(paragraphs, function (index, paragraph) {
                        if (paragraph.type == 'RichText') {
                            richTextParagraphs.push(paragraph);
                            paragraph.load("id,richText/text");
                        }
                    });
                    return context.sync();
                })
                .then(function () {
                    // Display all rich text paragraphs to the console
                    $.each(richTextParagraphs, function (index, richTextParagraph) {
                        var query = findQuery(richTextParagraph.richText.text);
                        console.log("Query : " + query);
                        fetchAndDisplay(context, richTextParagraph, query);
                    });
                    return context.sync();
                });
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

    }

})();

function findQuery(paragraph) {
    console.log("fqy". paragraph);
    var sentences = paragraph.split('.');
    var query = null
    for (var i = sentences.length - 1; i >= 0; i--) {
        var sentence = sentences[i];
        if (sentence.length > 0) {
            query = sentence;
            break;
        }
    }

    var cnt = 0;
    for (var i = query.length - 1; i >= 0; i--) {
        if (query[i] === " ") {
            cnt++;
        }
        if (cnt === 10 || i === 0) {
            return query.substring(i, query.length).trim();
        }
    }
}

function removeLastLine(text)
{
    var flag = false;
    for (var i = text.length - 1; i >= 0; i--) {
        if(text[i] != " " && text[i] != ".")
        {
            flag = true;
        }

        if (text[i] === "." && flag)
        {
            return text.substring(0,i+1);
        }

        if(i === 0)
        {
            return "";
        }
    }
}

function onItemClick(context, richTextParagraph, element)
{
    var text = removeLastLine(richTextParagraph.richText.text);
    richTextParagraph.insertHtmlAsSibling("After", text+' <font color="green"><b>'+element+'</b></font>');
    richTextParagraph.delete();

    context.sync();
}

function callback(context, richTextParagraph, data, query) {
    console.log('WE GOT RESULTS', data, status);
     $("#query").text('Searching for : ' + query);
     document.getElementById("onboard").style.display = "none";
    $("#results").empty();
    var cnt = 0;
    data.forEach(function (element) {
        cnt++;
        console.log('HERE ARE THE ELEMENTS', element);
        $("#results").append('<li id="id'+cnt+'">' + element + '</li>');
        document.getElementById("id"+cnt).addEventListener('click', function(){
            onItemClick(context, richTextParagraph, element);
        });
    }, this);
}

var URL = "https://jarvis-note1.azurewebsites.net/hello?q=";
var cache = {}
function fetchAndDisplay(context, richTextParagraph, query) {
    if (cache.hasOwnProperty(query)) {
        //callback(cache[query]);
        return;
    }
    else {
        cache = {}
    }

    var URLtoSend = URL + query;
    console.log("Getting results");
    $.get(URLtoSend).then(function(data, status)
    {
        if(status === "success")
        {
            cache[query] = data;
            callback(context, richTextParagraph, data, query);
        }
    });
}

