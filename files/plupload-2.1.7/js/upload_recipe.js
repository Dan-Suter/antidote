// Initialize the widget when the DOM is ready

$(function() {
	$("#uploader").plupload({
		// General settings
		runtimes : 'html5,flash,silverlight,html4',
		url : '../files/plupload-2.1.7/upload_recipe.php',
		// User can upload no more then 20 files in one go (sets multiple_queues to false)
		max_file_count: 1,
		multipart_params: {newfilename:$("#file_name").val()},
		chunk_size: '1mb',

		// Resize images on clientside if we can
		resize : {
			width : 300, 
			height : 300, 
			quality : 50,
			crop: true // crop to exact dimensions
		},
		
		filters : {
			// Maximum file size
			max_file_size : '10mb',
			// Specify what files to browse for
			mime_types: [
				{title : "Image files", extensions : "jpg,gif,png"}
			]
		},

		// Rename files by clicking on their titles
		rename: true,
		
		// Sort files
		sortable: true,

		// Enable ability to drag'n'drop files onto the widget (currently only HTML5 supports that)
		dragdrop: true,

		// Views to activate
		views: {
			list: true,
			thumbs: true, // Show thumbs
			active: 'thumbs'
		},

		// Flash settings
		flash_swf_url : '../../js/Moxie.swf',

		// Silverlight settings
// Silverlight settings
        silverlight_xap_url : '/plupload/js/Moxie.xap',
         
        // PreInit events, bound before any internal events
        preinit : {
            Init: function(up, info) {
                log('[Init]', 'Info:', info, 'Features:', up.features);
            },
 
            UploadFile: function(up, file) {
                log('[UploadFile]', file);
 
                // You can override settings before the file is uploaded
                // up.setOption('url', 'upload.php?id=' + file.id);
                // up.setOption('multipart_params', {param1 : 'value1', param2 : 'value2'});
            }
        },
 
        // Post init events, bound after the internal events
        init : {
            UploadComplete: function(up, files) {
                // Called when all files are either uploaded or failed
                log('[UploadComplete]');
                d = new Date();
				$("#update_img").attr("src", "/images/recipes/"+$("#file_name").val()+"?d="+d.getTime());
				$("#popup").bPopup().close();

            }
        }
    });
 
 
    function log() {
        var str = "";
 
        plupload.each(arguments, function(arg) {
            var row = "";
 
            if (typeof(arg) != "string") {
                plupload.each(arg, function(value, key) {
                    // Convert items in File objects to human readable form
                    if (arg instanceof plupload.File) {
                        // Convert status to human readable
                        switch (value) {
                            case plupload.QUEUED:
                                value = 'QUEUED';
                                break;
 
                            case plupload.UPLOADING:
                                value = 'UPLOADING';
                                break;
 
                            case plupload.FAILED:
                                value = 'FAILED';
                                break;
 
                            case plupload.DONE:
                                value = 'DONE';
                                break;
                        }
                    }
 
                    if (typeof(value) != "function") {
                        row += (row ? ', ' : '') + key + '=' + value;
                    }
                });
 
                str += row + " ";
            } else {
                str += arg + " ";
            }
        });
 
       
    }

	// Handle the case when form was submitted before uploading has finished
	$('#form').submit(function(e) {
		// Files in queue upload them first
		if ($('#uploader').plupload('getFiles').length > 0) {

			// When all files are uploaded submit form
			$('#uploader').on('complete', function() {
				$('#form')[0].submit();
			});

			$('#uploader').plupload('start');
		} else {
			alert("You must have at least one file in the queue.");
		}
		return false; // Keep the form from submitting
	});
});
