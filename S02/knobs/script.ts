Office.initialize = function (reason) 
{
	$(document).ready(bindEvents);
}

function bindEvents() : void
{
    $("#clone").click
    (
        function()
        {
            let p = $(".knob_panel")[0];
            let x = $(p).clone(true).appendTo("body");

            $(x).find(".v_knob")[0]['angle'] = 0;
            $(x).find(".s_knob")[0]['angle'] = 0;
            $(x).find(".s_knob")[0]['scale'] = 1;

            $(x).css( { 'visibility':'visible','display':'' } );

            getCell(x);
        }
    );

    $("#delete").click
    (
        function()
        {
            $(this).parent().parent().parent().remove();
        }
    )

    $(".v_knob")[0]['angle'] = 0;
    
    $(".v_knob").bind
    (
        'mousewheel DOMMouseScroll',

        function(event)
        {
            let parent = $(this).parent();
            let scale = parseFloat($(parent).find(".scale").text());
            let value = parseFloat($(this).find(".cell_value").text());

            if (event.originalEvent['wheelDelta'] > 0 || event.originalEvent['detail'] < 0)
            {
                this['angle'] += 3.6; // scroll up
                value = value + scale;
            }
            else
            {
                this['angle'] -= 3.6; // scroll down
                value = value - scale;
            }
            
            this['angle'] = this['angle'] % 360;
            
            $(this).find(".led").css
            (
                {
                '-moz-transform':'rotate(' + this['angle'] + 'deg)',
                '-webkit-transform':'rotate(' + this['angle'] + 'deg)',
                '-o-transform':'rotate(' + this['angle'] + 'deg)',
                '-ms-transform':'rotate(' + this['angle'] + 'deg)',
                'transform':'rotate(' + this['angle'] + 'deg)'
                }
            );

            $(this).find(".cell_value").text(value);

            setCell(this);

            event.preventDefault();
        }
    );

    $(".s_knob")[0]['angle'] = 0;
    $(".s_knob")[0]['scale'] = 1;

    $(".s_knob").bind
    (
        'mousewheel DOMMouseScroll',

        function(event)
        {
            if (event.originalEvent['wheelDelta'] > 0 || event.originalEvent['detail'] < 0)
            {
                this['angle'] += 20; // scroll up
                this['scale'] = this['scale'] * 10;
            }
            else
            {
                this['angle'] -= 20; // scroll down
                this['scale'] = this['scale'] / 10;
            }

            if (this['angle'] > 120) { this['angle'] =  120;}
            if (this['angle'] <-120) { this['angle'] = -120;}

            if (this['scale'] > 1000000) {this['scale'] = 1000000;}
            if (this['scale'] < 0.000001) {this['scale'] = 0.000001;}

            this['scale'] = this['scale'].toFixed(7);

            $(this).find(".led").css
            (
                {
                '-moz-transform':'rotate(' + this['angle'] + 'deg)',
                '-webkit-transform':'rotate(' + this['angle'] + 'deg)',
                '-o-transform':'rotate(' + this['angle'] + 'deg)',
                '-ms-transform':'rotate(' + this['angle'] + 'deg)',
                'transform':'rotate(' + this['angle'] + 'deg)'
                }
            );
            
            let v = this['scale'].toString();

            if (this['scale'] > 0.1)
            {
                $(this).find(".scale").text(v.substring(0,v.indexOf(".")));
            }
            else
            {
                $(this).find(".scale").text(v.substring(0,v.indexOf("1")+1));
            }

            event.preventDefault();
        }
    );
}

function getCell(panel) : void
{
    Excel.run
    (
        function (context) 
        {
            var range = context.workbook.getSelectedRange();
            
            range.load('address');
            range.load('values');

            return context.sync().then
            (
                function()
                {
                   let cell = range.address.toString();
                   cell = cell.substr(cell.indexOf("!") + 1);
                   if (cell.indexOf(":"))
                   {
                      cell = cell.substr(cell.indexOf(":") + 1);
                   }
                   $(panel).find(".cell_address").text(cell);
                   $(panel).find(".cell_value").text(0 + range.values[0][0]);
                }
            );
        }
    )
    .catch
    (
        function (error) { }
    );
}

function setCell(panel) : void
{
    Excel.run
    (
        function (context) 
        {
            let cell = $(panel).find(".cell_address").text();
            let sheet = context.workbook.worksheets.getActiveWorksheet();
            let range = sheet.getRange(cell);

            range.values = [[ $(panel).find(".cell_value").text() ]];

            return context.sync();
        }
    )
    .catch
    (
        function (error) { }
    );
}