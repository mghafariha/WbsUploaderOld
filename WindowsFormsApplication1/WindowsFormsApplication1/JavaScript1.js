function sortBy(field) {
    return function (a, b) {
        var x = parseInt(a[field].split("-")[0]),
            y = parseInt(a[field].split("-")[1]),
            z = parseInt(b[field].split("-")[0]),
            w = parseInt(b[field].split("-")[1])
        if (x < z)
            return -1;
        else if (x > z)
            return 1;
        else {
            if (y < w)
                return -1;
            else
                return 1;
        }
    };
}

var f=function (a, b) {
    var x=parseInt(a.split("-")[0]),
        y=parseInt(a.split("-")[1]),
        z=parseInt(b.split("-")[0]),
        w=parseInt(b.split("-")[1])
    if (x < z)
        return -1;
    else if (x > z)
        return 1;
    else
    {
        if (y < w)
            return -1;
        else
            return 1;
    }
   
}