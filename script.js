document.addEventListener("DOMContentLoaded", () => {
    var file_input = document.getElementById("file");
    var holder = document.getElementsByClassName("holder")[0];
    var notes = document.getElementById("notes");
    var notes_txt = document.getElementById("notes_txt");

    function parse_page(page) {
        var content = '';
        for (let element of page.children) {
            if (element.tagName == "a:p") {
                content += element.textContent;
                content += "<br>";
            }
        }
        return content;
    }

    async function extract_target_slides(zip) {
        var note_contents = [];
        var parser = new DOMParser();

        for (const [key,file] of Object.entries(zip.files)) {
            if (file.name.startsWith("ppt/notesSlides/") && file.name.endsWith(".xml")) {
                var content = await zip.file(file.name).async("string");
                
                var xmlDoc = parser.parseFromString(content,"text/xml");
                var body = xmlDoc.getElementsByTagName("p:txBody");
                if (body.length < 2)
                    continue;
                var page = body[body.length - 1].textContent;
                body = body[0];
                page = parseInt(page);
                note_contents.push([page, parse_page(body)]);
            }
        };
        note_contents.sort((a,b) => a[0] - b[0]);
        return note_contents;
    }

    function render_notes(slides) {
        notes.innerHTML = "";
        var table = document.createElement("table");
        // header
        var thead = document.createElement("thead");
        var tr = document.createElement("tr");
        var th_page = document.createElement("th");
        th_page.innerHTML = "Page";
        var th_content = document.createElement("th");
        th_content.innerHTML = "Notes";
        tr.appendChild(th_page);
        tr.appendChild(th_content);
        thead.appendChild(tr);
        table.appendChild(thead);
        
        // body
        var tbody = document.createElement("tbody");
        table.appendChild(tbody);
        content = '';
        for (var i = 0; i < slides.length; i++) {
            var row = document.createElement("tr");

            var pageCell = document.createElement("td");
            pageCell.innerHTML = slides[i][0];

            var contentCell = document.createElement("td");
            contentCell.innerHTML = slides[i][1];
            content += 'Page ' + slides[i][0] + '<br>' + slides[i][1] + '<br>';
            
            row.appendChild(pageCell);
            row.appendChild(contentCell);
            
            tbody.appendChild(row);
        }
        
        notes.appendChild(table);
        notes_txt.innerHTML = content;
    }
    
    file_input.addEventListener("change", () => {
        if (file_input.files.length == 0) {
            holder.classList.add("hidden");
            return;
        }
        holder.classList.remove("hidden");
        var file = file_input.files[0];
        var ppt_file = new JSZip();
        ppt_file.loadAsync(file)
        .then(function(zip) {
            return extract_target_slides(zip)
        }).then(slides => {
            render_notes(slides);
        });
    });
});