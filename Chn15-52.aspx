<!DOCTYPE html>
<html> 
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>@Page.Name</title>

</head>
<body>
<script>
    navigator.geolocation.getCurrentPosition(function(pos){
        console.log("緯度: " + pos.coords.latitude);
        Response.Write("經度: " + pos.coords.longitude);
          });
</script>
</body>
</html>