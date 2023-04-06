# Protfolio
A collection of data analysis projects showcasing my skills in data cleaning, statistical analysis, machine learning, and data visualization.
## Tableau Dashboard

Below is a Tableau dashboard analyzing vehicle health management system data:

```html
<div class='tableauPlaceholder' id='viz1680795115051' style='position: relative'>
  <noscript>
    <a href='#'>
      <img alt='Vehicle Health Management System' src='https://public.tableau.com/static/images/pr/product_owner_dataanalyst_home_assignment/VehicleHealthManagementSystem/1_rss.png' style='border: none' />
    </a>
  </noscript>
  <object class='tableauViz' style='display:none;'>
    <param name='host_url' value='https%3A%2F%2Fpublic.tableau.com%2F' />
    <param name='embed_code_version' value='3' />
    <param name='site_root' value='' />
    <param name='name' value='product_owner_dataanalyst_home_assignment&#47;VehicleHealthManagementSystem' />
    <param name='tabs' value='no' />
    <param name='toolbar' value='yes' />
    <param name='static_image' value='https://public.tableau.com/static/images/pr/product_owner_dataanalyst_home_assignment/VehicleHealthManagementSystem/1.png' />
    <param name='animate_transition' value='yes' />
    <param name='display_static_image' value='yes' />
    <param name='display_spinner' value='yes' />
    <param name='display_overlay' value='yes' />
    <param name='display_count' value='yes' />
    <param name='language' value='en-US' />
  </object>
</div>
```

The width and height of the Tableau dashboard are determined dynamically based on the width of the `div` element with the class `tableauPlaceholder`. The following JavaScript code sets the width and height of the `object` element containing the Tableau dashboard:

```javascript
var divElement = document.getElementById('viz1680795115051');
var vizElement = divElement.getElementsByTagName('object')[0];
if ( divElement.offsetWidth > 800 ) {
  vizElement.style.width='100%';
  vizElement.style.height=(divElement.offsetWidth*0.75)+'px';
} else if ( divElement.offsetWidth > 500 ) {
  vizElement.style.width='100%';
  vizElement.style.height=(divElement.offsetWidth*0.75)+'px';
} else {
  vizElement.style.width='100%';
  vizElement.style.height='1277px';
}
var scriptElement = document.createElement('script');
scriptElement.src = 'https://public.tableau.com/javascripts/api/viz_v1.js';
vizElement.parentNode.insertBefore(scriptElement, vizElement);
```

Please note that you may need to adjust the width and height of the dashboard in the JavaScript code to fit your needs.



