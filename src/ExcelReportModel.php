<?php

namespace customit\excelreport;

use Yii;
use Box\Spout\Common\Exception\IOException;
use Box\Spout\Common\Exception\UnsupportedTypeException;
use Box\Spout\Common\Type;
use Box\Spout\Writer\Style\BorderBuilder;
use Box\Spout\Writer\Style\StyleBuilder;
use Box\Spout\Writer\WriterFactory;
use Box\Spout\Writer\Style\Border;
use Box\Spout\Writer\Style\Color;
use yii\base\InvalidConfigException;
use box\spout;
use yii\queue\Queue;
use customit\excelreport\ExcelReportHelper;

class ExcelReportModel {

    /*
     * @var ActiveDataProvider
     */
    private $_provider;
    /*
     * @var array
     */
    private $_columns;
    /*
     * @var string
     */
    public $filename;
    /*
     * @var bool
     */
    public $stripHtml = true;
    /*
     * @var string
     */
    public $folder = '@app/runtime/export';
    /*
     * @var array
     */
    public $headerStyleOptions = [];
    /*
     * @var array
     */
    public $boxStyleOptions = [];
    /*
     * @var bool
     */
    public $enableFormatter = true;
    /*
     * @var array
     */
    public $styleOptions = [];

    /*
     * @var WriterInterface
     */
    protected $_objWriter;

    protected $_objWorksheet;
    /*
     * @var int
     */
    protected $_endCol = 1;
    /*
     * @var string
     */
    protected $_exportType = 'xlsx';
    /*
     * @var int
     */
    protected $_endRow = 0;
    /*
     * @var int
     */
    protected $_beginRow = 1;
    /*
     * @var Queue
     */
    protected $queue;
    /*
     * @var array
     */
    protected $_bodyData = [];

    /**
     * ExcelReportModel constructor.
     * @param string $columns base64 serialize array of gridview columns
     * @param Queue $queue current queue for status reports
     * @param string $fileName
     * @param string $dataProvider base64 serialize array of ActiveDataProvider
     * @param array $config
     */
    public function __construct($columns, $queue, $fileName, $dataProvider, array $config = [])
    {
        $this->_provider = ExcelReportHelper::reverseClosureDetect(unserialize(base64_decode($dataProvider)));
        $this->_columns = $this->cleanColumns(ExcelReportHelper::reverseClosureDetect(unserialize(base64_decode($columns))));
        $this->queue = $queue;
        $this->filename = $fileName;
    }

    /**
     * Remove extra columns
     * @param array $columns array of gridview columns
     * @return array
     */
    public function cleanColumns($columns) {
        foreach ($columns as $key => &$column) {
            if (!empty($column['hiddenFromExport'])) {
                unset($columns[$key]);
                continue;
            }            
            
            if (isset($column['class']) && $column['class'] == 'yii\\grid\\ActionColumn') {
                unset($columns[$key]);
            } elseif (isset($column['class']) && $column['class'] == 'yii\\grid\\SerialColumn') {
                unset($columns[$key]);
            } elseif (isset($column['value'])) {
                $column['attribute'] = $column['value'];
            }
            
        }
        return $columns;
    }

    /**
     * Entry point
     */
    public function start() {
        $config = [
            'extension' => 'xlsx',
            'writer' => Type::XLSX,
        ];
        $this->initExport();
        try {
            $this->initExcelWriter($config);
        } catch (InvalidConfigException $e) {
            Yii::error($e->getMessage());
            exit;
        } catch (IOException $e) {
            Yii::error($e->getMessage());
            exit;
        } catch (UnsupportedTypeException $e) {
            Yii::error($e->getMessage());
            exit;
        }
        $this->initExcelWorksheet();
        $this->generateHeader();
        $totalCount = $this->generateBody();
        //Write data to file
        $this->_objWriter->close();
        $this->queue->setProgress($this->_endRow, $totalCount);
        //Unset vars
        $this->cleanup();
    }

    /**
     * Initializes export settings
     */
    public function initExport()
    {
        $this->setDefaultStyles('header');
        $this->setDefaultStyles('box');

        if (!isset($this->filename)) {
            $this->filename = 'grid-export';
        }
    }

    /**
     * Appends slash to path if it does not exist
     *
     * @param string $path
     * @param string $s the path separator
     *
     * @return string
     */
    public static function slash($path, $s = DIRECTORY_SEPARATOR)
    {
        $path = trim($path);
        if (substr($path, -1) !== $s) {
            $path .= $s;
        }
        return $path;
    }

    /**
     * Sets default styles
     *
     * @param string $section
     */
    protected function setDefaultStyles($section)
    {
        $defaultStyle = [];
        $opts = '';
        if ($section === 'header') {
            $opts = 'headerStyleOptions';

            $border = (new BorderBuilder())
                ->setBorderBottom(Color::BLACK, Border::WIDTH_MEDIUM, Border::STYLE_SOLID)
                ->build();
            $defaultStyle = (new StyleBuilder())
                ->setFontBold()
                ->setBackgroundColor('FFE5E5E5')
                ->setBorder($border)
                ->build();

        } elseif ($section === 'box') {
            $opts = 'boxStyleOptions';

            $border = (new BorderBuilder())
                ->setBorderBottom(Color::BLACK, Border::WIDTH_MEDIUM, Border::STYLE_SOLID)
                ->build();
            $defaultStyle = (new StyleBuilder())
                ->setBorder($border)
                ->build();
        }
        if (empty($opts)) {
            return;
        }

        $this->$opts = $defaultStyle;
    }

    /**
     * Initializes Spout Writer Object Instance
     *
     * @param array $config
     * @throws InvalidConfigException
     * @throws \Box\Spout\Common\Exception\UnsupportedTypeException
     * @throws \Box\Spout\Common\Exception\IOException
     */
    public function initExcelWriter($config)
    {
        $this->folder = trim(Yii::getAlias($this->folder));
        $file = self::slash($this->folder) . $this->filename . '.' . $config['extension'];
        if (!file_exists($this->folder) && !mkdir($this->folder, 0777, true)) {
            throw new InvalidConfigException(
                "Invalid permissions to write to '{$this->folder}' as set in `Export::folder` property."
            );
        }

        $this->_objWriter = WriterFactory::create($config['writer']);
        $this->_objWriter->setShouldUseInlineStrings(true);
        $this->_objWriter->setShouldUseInlineStrings(false);

        $this->_objWriter->openToFile($file);
    }

    /**
     * Get Worksheet Instance
     */
    public function initExcelWorksheet()
    {
        $this->_objWorksheet = $this->_objWriter->getCurrentSheet();
        $this->_objWorksheet->setName(Yii::t('customit','Report'));
    }

    /**
     * Generates the output data header content.
     */
    public function generateHeader()
    {
        if (count($this->_columns) == 0) {
            return;
        }
        $styleOpts = $this->headerStyleOptions;
        $headValues = [];
        $this->_endCol = 0;
        //Generate labels array
        foreach ($this->_columns as $column) {
            $this->_endCol++;
            if (isset($column['label'])) {
                $head =  $column['label'];
            } elseif (isset($column['header'])) {
                $head =  $column['header'];
            } else {
                $head = '#';
            }
            $headValues[] = $head;
        }
        //Write header content
        $this->setRowValues($headValues, $styleOpts);
    }

    /**
     * Sets the values of excel row
     *
     * @param array $values
     * @param array $style
     *
     */
    protected function setRowValues($values, $style = null)
    {
        if ($this->stripHtml) {
            array_walk_recursive($values, function (&$item, $key) {
                $item = strip_tags($item);
            });
        }
        if (!empty($style)) {
            $this->_objWriter->addRowWithStyle($values, $style);
        } else {
            $this->_objWriter->addRows($values);
        }
    }

    /**
     * Generates the output data body content.
     *
     * @return integer the number of output rows.
     */
    public function generateBody()
    {
        $this->_endRow = 0;
        $totalCount = $this->_provider->getTotalCount();        
        
        foreach ($this->_provider->query->each() as $value) {  
            $this->generateRow($value);
            $this->_endRow++;
            //Change queue process progress                
            if (($this->_endRow % 1000) == 0) $this->queue->setProgress($this->_endRow-1, $totalCount);
        }

        $this->setRowValues($this->_bodyData);
        return $totalCount;
    }

    /**
     * Generates an output data row with the given data.
     *
     * @param mixed $data the data model to be rendered
     */
    public function generateRow($data)
    {
        $this->_endCol = 0;
        $key = count($this->_bodyData);

        foreach ($this->_columns as $column) {
            $var = isset($column['attribute']) ? $column['attribute'] : null;
            if (is_string($var)) {
                $valueChain = explode('.', $var);
                $bufObj = $data;
                if (count($valueChain) > 1) {
                    foreach ($valueChain as $vc) {
                        $bufObj = is_object($bufObj) ? $bufObj->$vc : "---";
                    }
                    $value = $bufObj;
                } else {
                    $value = is_object($data) ? $data->$var : "---";
                }
            } elseif (is_object($var) && ExcelReportHelper::is_closure($var)) {
                $value = call_user_func($var, $data);
            } else {
                $value = null;
            }
            $this->_bodyData[$key][] = isset($column['format']) ? Yii::$app->formatter->format($value, $column['format']) : $value;
        }
    }

    /**
     * Cleans up current objects instance
     */
    protected function cleanup()
    {
        unset($this->_provider, $this->_objWriter);
    }
}
