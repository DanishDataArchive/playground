var strFormId = "#formkipfolder";

$(function() {
	
	
	$('#ddlCounty').on('change', function(){
		$('#ddlParish').html('');
		$.post("soegpersonenkelajax.asp?action=getHerreds", 
				$(strFormId).serialize(),
				function(data) {
					$('#ddlHerred').html(data);
				}
		);
		$.post("soegpersonenkelajax.asp?action=getParishes", 
				$(strFormId).serialize(),
				function(data) {
					$('#ddlParish').html(data);
				}
		);
	});
	$('#ddlHerred').on('change', function(){
		$.post("soegpersonenkelajax.asp?action=getParishes", 
				$(strFormId).serialize(),
				function(data) {
					$('#ddlParish').html(data);
				}
		);
	});
	$('#btnSearch').on('click', function() {
		if($('#ddlCounty').val() == "alle")
			alert("Venligst vælg Amt");
		else
			performSearch();
	});
});

function performSearch() {
	$('#searchResults').html('<img src="loading.gif" alt="Loading..." />');
	$.post("soegpersonenkelajax.asp?action=search", 
			$(strFormId).serialize(),
			function(data) {
				$('#searchResults').html(data);
			}
	);

}