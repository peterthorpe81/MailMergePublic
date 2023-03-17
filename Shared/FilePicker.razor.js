export function filepicker(clientId, accessToken, dotNetObjRef) {
    // public interface
    const proxy = {
        launchOneDrivePicker
    };

    function launchOneDrivePicker() {
        var odOptions = {
            clientId: clientId,
            action: "query",
            multiSelect: false,
            advanced: {
                filter: ".xlsx",
                queryParameters: "select=id",
                accessToken: accessToken
            },
            success: function (response) {
                if (response && response.value.length > 0) {
                    dotNetObjRef.invokeMethodAsync('FileSelected', response.value[0].parentReference.driveId, response.value[0].id)
                        .then(data => {
                            console.log(data);
                        });
                    };
            },
            cancel: function () { },
            error: function (e) { alert('An error occourred: ' + e); }
        }
        OneDrive.open(odOptions);
    }

    return proxy;
  }

