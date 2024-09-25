document.addEventListener("DOMContentLoaded", () => {
    var file_input = document.getElementById("file");
    var holder = document.getElementsByClassName("holder")[0];
    var empty = document.getElementById('empty');
    var notes = document.getElementById("notes");
    var notes_txt = document.getElementById("notes_txt");

    async function extract_target_slides(zip) {
        var note_contents = [];
        var parser = new DOMParser();

        for (const [key,file] of Object.entries(zip.files)) {
            if (file.name.startsWith("ppt/notesSlides/") && file.name.endsWith(".xml")) {
                var content = await zip.file(file.name).async("string");
                
                var xmlDoc = parser.parseFromString(content,"text/xml");
                var page = new Page(xmlDoc);
                if (page.sldNum >= 0) {
                    note_contents.push(page);
                }
            }
        };
        note_contents.sort((a,b) => a.sldNum - b.sldNum);
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
            pageCell.innerHTML = slides[i].sldNum;

            var contentCell = document.createElement("td");
            contentCell.innerHTML = slides[i].getContent();
            if (contentCell.innerText.length == 0)
                continue;
            content += 'Page ' + slides[i].sldNum + '<br>' + slides[i].getRawContent() + '<br>';
            
            row.appendChild(pageCell);
            row.appendChild(contentCell);
            
            tbody.appendChild(row);
        }
        
        notes.appendChild(table);
        notes_txt.innerHTML = content;
        if (content.length == 0) {
            holder.classList.add("hidden");
            empty.classList.remove("hidden");
        }
    }
    
    file_input.addEventListener("change", () => {
        empty.classList.add("hidden");
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

    var toggle = document.getElementById("toggle");
    var header = document.getElementById("header");
    var footer = document.getElementsByTagName("footer")[0];
    var table_holder = document.getElementById("table_holder");
    var txt_holder = document.getElementById("txt_holder");
    var mode = 0;
    toggle.addEventListener("click", () => {
        mode = (mode + 1) % 3;
        if (mode == 0) {
            header.classList.remove("hidden");
            footer.classList.remove("hidden");
            table_holder.classList.remove("hidden");
            txt_holder.children[0].classList.remove("hidden");
        } else if (mode == 1) {
            header.classList.add("hidden");
            footer.classList.add("hidden");
            txt_holder.classList.add("hidden");
            table_holder.children[0].classList.add("hidden");
        } else if (mode == 2) {
            txt_holder.classList.remove("hidden");
            txt_holder.children[0].classList.add("hidden");
            table_holder.classList.add("hidden");
            table_holder.children[0].classList.remove("hidden");
        }    
    });
});