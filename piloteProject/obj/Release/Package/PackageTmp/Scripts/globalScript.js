

function addProject() {
    $.ajax({
        type: "GET",
        url: 'Home/AddProjectFormulaire',
        success: function (msg) {
            $("#getModal").empty();
            $("#getModal").html(msg);
            $("#ptProjectModal").modal("show");
        }
    })
}

function showProject(id) {

    $.ajax({
        type: "GET",
        url: 'Home/AddProjectFormulaire?Id=' + id,
        success: function (msg) {
            $("#getModal").empty();
            $("#getModal").html(msg);
            $("#ptProjectModal").modal("show");
        }
    })
}

function showProjectDecided(id) {
    var d = "true";
    $.ajax({
        type: "GET",
        url: 'AddProjectFormulaire?Id=' + id + '&decide=' + d,
        success: function (msg) {
            $("#getModal").empty();
            $("#getModal").html(msg);
            $("#ptProjectModal").modal("show");
        }
    })
}



function noteProject(id) {
    $.ajax({
        type: "GET",
        url: 'Notes/NoteProject?ProjectId=' + id,
        success: function (msg) {
            $("#getModal").empty();
            $("#getModal").html(msg);
            $("#ptNoteProjectModal").modal("show");
        }
    })
}

function showNoteProject(noteId, projectId) {
    $.ajax({
        type: "GET",
        url: 'Notes/NoteProject?NoteId=' + noteId + '&ProjectId=' + projectId,
        success: function (msg) {
            $("#getModal").empty();
            $("#getModal").html(msg);
            $("#ptNoteProjectModal").modal("show");
        }
    })
}

function decision(Id) {
    $.ajax({
        type: "GET",
        url: 'Home/ProjectDecision?Id=' + Id,
        success: function (msg) {
            $("#getModal").empty();
            $("#getModal").html(msg);
            $("#ptDecisionProjectModal").modal("show");

        }
    })
}





function showInput(that) {
    if (that.value == "Autre") {
        document.getElementById("inputShowHide").style.display = "block";
    } else {
        document.getElementById("inputShowHide").style.display = "none";
    }
}

function Delete(id) {
    swal({
        title: "Vous êtes sûr ?",
        text: "",
        icon: "warning",
        buttons: true,
        dangerMode: true,
    }).then((willDelete) => {
        if (willDelete) {
            swal("Supprimer !", {
                icon: "success",
            });
            $.ajax({
                type: "GET",
                url: 'Home/ProjectDelete?Id=' + id,
                success: function (msg) {
                    window.location.reload(true);
                }
            });


        }


    });

}

function showAlertConfirmation(that, id) {
    swal("Succès", "Le statut à été modifié!", "success");
    $.ajax({
        type: "GET",
        url: '../Home/ProjectStatus?Id=' + id + '&status=' + that.value,
    })
}

function showAlertConfirmationMeteo(that, id) {
    swal("Succès", "Le météo à été modifié!", "success");
    $.ajax({
        type: "GET",
        url: '../Home/ProjectMeteo?Id=' + id + '&meteo=' + that.value,
    })
}

function emptyConsistance(that) {
    if (that.value != "Autre") {
        $('#Consistance').val('');
    }
}

function validateForm(event) {
    var services = '';
    var direction = '';
    var initiative = '';
    $('#Services > option:selected').each(function () {
        services = services + $(this).val() + ';';
    });
    $('#Directions > option:selected').each(function () {
        direction = direction + $(this).val() + ';';
    });
    $('#Initiative > option:selected').each(function () {
        initiative = initiative + $(this).val() + ';';
    });
    $("#DirectionsHidden").val(direction);
    $("#ServicesHidden").val(services);
    $("#InitiativeHidden").val(initiative);
}

function showService()
{
    var tableIds = new Array();
    $('#Directions > option:selected').each(function () {
        tableIds.push($(this).attr("id"));
    });

    $('#Services > option').each(function () {
        var idDirection = $(this).attr("class");
        if (tableIds.indexOf(idDirection) == -1) {
            $(this).hide();
        }
        else {
            $(this).show();
        }

    });
    //sort
    var select = $('#Services');
    select.html(select.find('option').sort(function (x, y) {
        // to change to descending order switch "<" for ">"
        return $(x).text() > $(y).text() ? 1 : -1;
    }));
    
}
