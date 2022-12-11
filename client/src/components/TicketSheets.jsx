import {useEffect, useState} from 'react';
import axios from 'axios';

const TicketSheets = () => {
    const [ticketSheetNames, setTicketSheetNames] = useState([]);
    useEffect(() => {
        const url = `${process.env.REACT_APP_HOST}/api/ticket_sheet`;
        const config = {};
        axios.get(url, config)
            .then((res) => {
                setTicketSheetNames(res.data);
            })
            .catch((err) => {
                console.log(err);
            });
    }, []);

    const handleInsertion = (name, event) => {
        const url = `${process.env.REACT_APP_HOST}/api/ticket_sheet/${name}`;
        //const config = {};
        axios.patch(url)
            .then((res) => {
                console.log(res);
            })
            .catch((err) => {
                console.log(err);
            });
    };

    return (
        <div>
            {ticketSheetNames.map((fileName, i) =>
                <div key={i}>
                    {fileName}
                    <button onClick={e => handleInsertion(fileName, e)}>Insert To DB</button>
                    <button onClick={e => {}}>Delete</button>
                </div>
            )}
        </div>
    );
};

export default TicketSheets;