<p align="center">
    <a href="https://custom-it.ru" target="_blank">
        <img src="https://avatars1.githubusercontent.com/u/31646762?s=200&v=4" height="100px">
    </a>
    <h1 align="center">Yii2 ExcelReport Extension</h1>
    <br>
</p>


An extension for generate excel file from GridView content

[![Latest Stable Version](https://poser.pugx.org/custom-it/yii2-excel-report/v/stable.svg)](https://packagist.org/packages/custom-it/yii2-excel-report)
[![Total Downloads](https://poser.pugx.org/custom-it/yii2-excel-report/downloads.svg)](https://packagist.org/packages/custom-it/yii2-excel-report)

Installation
------------

The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

Either run

```
php composer require --prefer-dist custom-it/yii2-excel-report
```

or add

```
"custom-it/yii2-excel-report": "*"
```

to the require section of your `composer.json` file.


Usage
-----

Once the extension is installed, simply use it in your code by  :

```php
echo \customit\excelreport\ExcelReport::widget([
    'columns' => $gridColumns,
    'stripHtml' => false,
    'searchClass' => get_class($searchModel),
    'searchMethod' => 'search',
]);
```
