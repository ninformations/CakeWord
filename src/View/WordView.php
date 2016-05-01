<?php
namespace CakeWord\View;

use Cake\Core\Exception\Exception;
use Cake\Event\EventManager;
use Cake\Network\Request;
use Cake\Network\Response;
use Cake\Utility\Inflector;
use Cake\View\View;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\PhpWord;

/**
 * @package  Cake.View
 */
class WordView extends View
{

    /**
     * @var PHPWord
     */
    private $phpWord;

    private $filename;

    /**
     * SubDir Name
     * @var string
     */
    public $subDir = 'docx';

    /**
     * Constructor
     *
     * @param \Cake\Network\Request $request Request instance.
     * @param \Cake\Network\Response $response Response instance.
     * @param \Cake\Event\EventManager $eventManager Event manager instance.
     * @param array $viewOptions View options. See View::$_passedVars for list of
     *   options which get set as class properties.
     *
     * @throws \Cake\Core\Exception\Exception
     */
    public function __construct(Request $request = null, Response $response = null, EventManager $eventManager = null, array $viewOptions = [])
    {
        parent::__construct($request, $response, $eventManager, $viewOptions);

        if (isset($viewOptions['name']) && $viewOptions['name'] == 'Error') {
            $this->subDir = null;
            $this->layoutPath = null;
            $response->type('html');
            return;
        }

        $this->phpWord = new PHPWord();
    }

    /**
     * Render method
     * @param  string $action - action to render
     * @param  string $layout - layout to use
     * @return string - rendered content
     */
    public function render($view = null, $layout = null)
    {
        $content = parent::render($view, $layout);
        if ($this->response->type() == 'text/html') {
            return $content;
        }

        $content = $this->getContent();
        $this->Blocks->set('content', $content);
        $this->response->download($this->getFilename());
        return $this->Blocks->get('content');
    }

    /**
     * Sets the filename
     * @param string $filename the filename
     * @return void
     */
    public function setFilename($filename)
    {
        $this->filename = $filename;
    }

    /**
     * Gets the filename
     * @return string filename
     */
    public function getFilename()
    {
        if (!empty($this->filename)) {
            return $this->filename . '.xlsx';
        }

        return Inflector::slug(str_replace('.docx', '', $this->request->url)) . '.doc';
    }

    private function getContent()
    {
        ob_start();

        $objWriter = IOFactory::createWriter($this->phpWord, 'Word2007');
        $objWriter->save('php://output');

        return ob_get_clean();
    }

    /**
     * @return PhpWord
     */
    public function getPhpWord()
    {
        return $this->phpWord;
    }

    /**
     * @param PhpWord $phpWord
     */
    public function setPhpWord($phpWord)
    {
        $this->phpWord = $phpWord;
    }

    /**
     * @return string
     */
    public function getSubDir()
    {
        return $this->subDir;
    }

    /**
     * @param string $subDir
     */
    public function setSubDir($subDir)
    {
        $this->subDir = $subDir;
    }
}
