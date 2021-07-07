function getUserGists(user, callback) {
  var requestUrl = "https://api.github.com/users/" + user + "/gists";

  $.ajax({
    url: requestUrl,
    dataType: "json",
  })
    .done(function (gists) {
      callback(gists);
    })
    .fail(function (error) {
      callback(null, error);
    });
}

async function getGraphData() {
  alert("exception", exception);

  try {
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken();
    console.log("bootstrapToken is", bootstrapToken);

    // The /api/DoSomething controller will make the token exchange and use the
    // access token it gets back to make the call to MS Graph.
    // getData("/api/DoSomething", bootstrapToken);
  } catch (exception) {
    if (exception.code === 13003) {
      // SSO is not supported for domain user accounts, only
      // Microsoft 365 Education or work account, or a Microsoft account.
    } else {
      // Handle error
    }
  }
}

function buildGistList(parent, gists, clickFunc) {
  gists.forEach(function (gist) {
    var listItem = $("<div/>").appendTo(parent);

    var radioItem = $("<input>")
      .addClass("ms-ListItem")
      .addClass("is-selectable")
      .attr("type", "radio")
      .attr("name", "gists")
      .attr("tabindex", 0)
      .val(gist.id)
      .appendTo(listItem);

    var desc = $("<span/>").addClass("ms-ListItem-primaryText").text(gist.description).appendTo(listItem);

    var desc = $("<span/>")
      .addClass("ms-ListItem-secondaryText")
      .text(" - " + buildFileList(gist.files))
      .appendTo(listItem);

    var updated = new Date(gist.updated_at);

    var desc = $("<span/>")
      .addClass("ms-ListItem-tertiaryText")
      .text(" - Last updated " + updated.toLocaleString())
      .appendTo(listItem);

    listItem.on("click", clickFunc);
  });
}

function buildFileList(files) {
  var fileList = "";

  for (var file in files) {
    if (files.hasOwnProperty(file)) {
      if (fileList.length > 0) {
        fileList = fileList + ", ";
      }

      fileList = fileList + files[file].filename + " (" + files[file].language + ")";
    }
  }

  return fileList;
}
function getGist(gistId, callback) {
  var requestUrl = "https://api.github.com/gists/" + gistId;

  $.ajax({
    url: requestUrl,
    dataType: "json",
  })
    .done(function (gist) {
      callback(gist);
    })
    .fail(function (error) {
      callback(null, error);
    });
}

function buildBodyContent(gist, callback) {
  // Find the first non-truncated file in the gist
  // and use it.
  for (var filename in gist.files) {
    if (gist.files.hasOwnProperty(filename)) {
      var file = gist.files[filename];
      if (!file.truncated) {
        // We have a winner.
        switch (file.language) {
          case "HTML":
            // Insert as-is.
            callback(file.content);
            break;
          case "Markdown":
            // Convert Markdown to HTML.
            var converter = new showdown.Converter();
            var html = converter.makeHtml(file.content);
            callback(html);
            break;
          default:
            // Insert contents as a <code> block.
            var codeBlock = "<pre><code>";
            codeBlock = codeBlock + file.content;
            codeBlock = codeBlock + "</code></pre>";
            callback(codeBlock);
        }
        return;
      }
    }
  }
  callback(null, "No suitable file found in the gist");
}
