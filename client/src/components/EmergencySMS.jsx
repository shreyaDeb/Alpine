const SMS = () => {

    return (
        <form action="/url" method="GET">
        <p>Type your message and hit submit to send a SMS:</p>
        <textarea name="message" id="userInputSMS"></textarea>
        <input type="button" value="Submit" class="submitBtn" onClick={() => {
        /*
            var AWS = require('aws-sdk');
            AWS.config.update({
                accessKeyId: '',
                secretAccessKey: '',
                region: 'us-west-1'
            })

            var params = {
                Message: 'TEST', // required
                PhoneNumber: '4066907119',
            };

            var publishTextPromise = new AWS.SNS({apiVersion: '2010-03-31'}).publish(params).promise();

            // Handle promise's fulfilled/rejected states
            publishTextPromise.then(
            function(data) {
                console.log("MessageID is " + data.MessageId);
            }).catch(
            function(err) {
                console.error(err, err.stack);
            });
	*/
        }}/>
        </form>
    )
  }
  
  export default SMS;
  