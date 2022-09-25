# xlsx-creator
You pass the Excel file path and json format data you want to input to this program.  
Make Excel with data you want to input and return base64 encoded string of it.

## example of data(json)

````
"cell": {
    "name1": {
        "value": "foo"
    }
    "name2": {
        "value": "bar"
    }
    "name3": {
        "value": 123
    }
},
"row": {
    "details1": {
        "insertRow": true
        "value": [
            {
                "name1-1": {
                    "value": "foo1-1"
                }
                "name1-2": {
                    "value": "bar1-1"
                }
                "name1-3": {
                    "value": 1231-1
                }
            },
            {
                "name2-1": {
                    "value": "foo2-1"
                }
                "name2-2": {
                    "value": "bar2-1"
                }
                "name2-3": {
                    "value": 1232-1
                }
            }

        ]
    },
    "details2": {
        "insertRow": false
        "value": [
            {
                "name1-1": {
                    "value": "foo1-1"
                }
                "name1-2": {
                    "value": "bar1-1"
                }
                "name1-3": {
                    "value": 1231-1
                }
            },
            {
                "name2-1": {
                    "value": "foo2-1"
                }
                "name2-2": {
                    "value": "bar2-1"
                }
                "name2-3": {
                    "value": 1232-1
                }
            }

        ]
    }
}
````