export function richedit(immediateUpdate = false) {
    // public interface
    const proxy = {
        init,
        destroy,
        update,
        insertText
    };

    let text_changed = false;
    let ckEditor;

    function init(element, dotNetReference) {
        ClassicEditor
            .create(element)
            .then(editor => {
                ckEditor = editor;
                editor.model.document.on('change:data', () => { 
                    if (immediateUpdate)
                        updateBlazor(editor, dotNetReference);
                    else
                        text_changed = true; 
                    
                });
                editor.ui.focusTracker.on('change:isFocused', (evt, name, isFocused) => {
                    if (!isFocused && text_changed && !immediateUpdate) {
                        updateBlazor(editor, dotNetReference);
                    }
                });
            })
            .catch(error => console.error(error));
    }

    function updateBlazor(editor, dotNetReference)
    {
        let data = editor.getData();

        const el = document.createElement('div');
        el.innerHTML = data;
        if (el.innerText.trim() == '')
            data = null;

        dotNetReference.invokeMethodAsync('EditorDataChanged', data);
        text_changed = false;
    }

    function update(editorValue) {
        if (ckEditor) {
            ckEditor.setData(editorValue);
        }
    }

    function destroy() {
        if (ckEditor) {
            ckEditor.destroy()
            .then(() => ckEditor = undefined)
            .catch(error => console.log(error));
        }
    }

    function insertText(text) {
        if (ckEditor) {
            ckEditor.model.change( writer => {
                writer.insertText(text, ckEditor.model.document.selection.getFirstPosition() );
            } );
            ckEditor.editing.view.focus();
        }
    }

    return proxy;
  }

