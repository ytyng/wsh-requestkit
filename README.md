* WSH RequestKit

<job>
<script language="JScript" src="wsf-request-toolkit.js"></script>
<script language="JScript">
var credential = RequestKit.getJson('https://example.com/get-credential');
var ie = new RequestKit.IE();
ie.navigate('http://example.com/login');
ie.login(credential.data.user_id, credential.data.password);
</script>
</job>
