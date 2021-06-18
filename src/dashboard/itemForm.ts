import { Components, Helper, SPTypes } from "gd-sprest-bs";
import { CanvasForm } from "../common";
import Strings from "../strings";

/**
 * Item Form
 */
class _ItemForm {
    private _displayForm: Components.IListFormDisplay = null;
    private _editForm: Components.IListFormEdit = null;
    private _updateEvent: Function = null;

    // The current form being displayed
    get form(): Components.IListFormDisplay | Components.IListFormEdit { return this._displayForm || this._editForm; }

    /** Public Methods */

    // Creates a new task
    create(onUpdate?: Function) {
        // Set the update event
        this._updateEvent = onUpdate;

        // Load the item
        this.load(SPTypes.ControlMode.New);
    }

    // Edits a task
    edit(itemId: number, onUpdate: () => void) {
        // Set the update event
        this._updateEvent = onUpdate;

        // Load the form
        this.load(SPTypes.ControlMode.Edit, itemId);
    }

    // Views the task
    view(itemId: number, onUpdate?: () => void) {
        // Set the update event
        this._updateEvent = onUpdate;

        // Load the form
        this.load(SPTypes.ControlMode.Display, itemId);
    }

    /** Private Methods */

    // Load the form information
    private load(mode: number, itemId?: number) {
        // Clear the forms
        this._displayForm = null;
        this._editForm = null;

        // Show a loading dialog
        Helper.SP.ModalDialog.showWaitScreenWithNoClose("Loading the Item").then(dlg => {
            // Load the form information
            Helper.ListForm.create({
                listName: Strings.Lists.FAQ,
                itemId
            }).then(info => {
                // Set the header
                CanvasForm.setHeader("<h5>Submit a Question</h5>");

                // Render the form based on the type
                if (mode == SPTypes.ControlMode.Display) {
                    // Render the display form
                    this._displayForm = Components.ListForm.renderDisplayForm({
                        info,
                        rowClassName: "mb-3"
                    });

                    // Update the body
                    CanvasForm.setBody(this._displayForm.el);
                } else {
                    let isNew = mode == SPTypes.ControlMode.New;
                    let el = document.createElement("div");

                    // Render the edit form
                    this._editForm = Components.ListForm.renderEditForm({
                        el,
                        info,
                        rowClassName: "mb-3",
                        controlMode: isNew ? SPTypes.ControlMode.New : SPTypes.ControlMode.Edit,
                        excludeFields: [
                            "Answer",
                            "Approved"
                        ],
                        onControlRendering: (ctrl, fld) => {
                            // See if this is the title field
                            if (fld.InternalName == "Title") {
                                // Make this a text area field
                                ctrl.type = Components.FormControlTypes.TextArea;

                                // Add validation
                                ctrl.onValidate = (ctrl, result) => {
                                    // Ensure this is less than 255 characters
                                    if (result.value && result.value.length > 255) {
                                        // Set the result
                                        result.isValid = false;
                                        result.invalidMessage = "The question must be <255 characters";
                                    }

                                    // Return the result
                                    return result;
                                }
                            }
                        }
                    });

                    // Render the save button
                    let elButton = document.createElement("div");
                    elButton.classList.add("float-end");
                    elButton.classList.add("mt-3");
                    el.appendChild(elButton);
                    Components.Button({
                        el: elButton,
                        text: isNew ? "Create" : "Update",
                        type: Components.ButtonTypes.OutlineSuccess,
                        onClick: () => { this.save(this._editForm, isNew); }
                    });

                    // Update the body
                    CanvasForm.setBody(el);
                }

                // Close the dialog
                dlg.close();

                // Show the form
                CanvasForm.show();
            });
        });
    }

    // Saves the form
    private save(form: Components.IListFormEdit, isNew: boolean) {
        // Validate the form
        if (form.isValid()) {
            // Display a loading dialog
            Helper.SP.ModalDialog.showWaitScreenWithNoClose("Validating the Form").then(dlg => {
                // Update the title
                dlg.setTitle((isNew ? "Creating" : "Updating") + " the Item");

                // Save the item
                form.save().then(item => {
                    // Call the update event
                    this._updateEvent ? this._updateEvent() : null;

                    // Close the dialogs
                    dlg.close();
                    CanvasForm.hide();
                });
            });
        }
    }
}
export const ItemForm = new _ItemForm();