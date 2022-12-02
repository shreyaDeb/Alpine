import { useRef, useState } from 'react';
import axios from "axios";
import { useDropzone } from 'react-dropzone';

// TODO: Drag and drop functionality for component
// TODO: Spinner for user waiting for file to upload
// TODO: Error messages for the user (i.e. Too many files, File too large)

const FileInput = () => {
    const fileInputRef = useRef(null);
    const [fileUploaded, setFileUploaded] = useState(false);
    //const [fileUploadFailure, setFileUploadFailure] = useState(false);
    //const {getRootProps, getInputProps} = useDropzone();

    const handleSubmit = async (event) => {
        event.preventDefault();
        if (!fileInputRef.current) {
            //  TODO: add a tag that says that you can't upload with out any files selected
            return;
        }
        const formData = new FormData();
        formData.append("ticket_sheet", fileInputRef.current.files[0]);
        const url = `${process.env.REACT_APP_HOST}/api/ticket_sheet`;
        const config = {
            headers: {
                'content-type': 'multipart/form-data',
            }
        };
        setFileUploaded(true);
        await axios.post(url, formData, config)
        fileInputRef.current.value = null;
    };

    return (
        <form encType="multipart/form-data" onSubmit={handleSubmit}>
            <label>Upload ticket spreadsheet:</label>
            <input
                type="file"
                ref={fileInputRef}
                accept=".xlsx,.csv"
            />
            <input type="submit" value="Upload File(s)" disabled={fileUploaded}/>
            {fileUploaded && (<p>Uploaded!</p>)}
        </form>
    );
};

export default FileInput;