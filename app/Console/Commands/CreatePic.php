<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use Intervention\Image\Facades\Image;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

class CreatePic extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'create:goods-pic';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'create goods pic';

    /**
     * @var string
     */
    private $dirRoot;

    /**
     * @var string
     */
    private $dirExcel;

    /**
     * @var string
     */
    private $dirIcon;

    /**
     * @var string
     */
    private $dirProduct;

    /**
     * @var string
     */
    private $dirFinal;

    /**
     * @var string
     */
    private $dirTtf;

    /**
     * Create a new command instance.
     *
     * @return void
     */
    public function __construct()
    {
        $this->dirRoot = storage_path() . '/goods/';
        $this->dirExcel = $this->dirRoot . 'excel/';
        $this->dirIcon = $this->dirRoot . 'icon/';
        $this->dirProduct = $this->dirRoot . 'product/';
        $this->dirFinal = $this->dirRoot . 'final/';
        $this->dirTtf = $this->dirRoot . 'ttf/';

        parent::__construct();
    }

    /**
     * Execute the console command.
     *
     * @return int
     */
    public function handle()
    {
        // 读取Excel信息
        $excelList = $this->readExcel();

        // 创建图片
        foreach ($excelList as $item) {
            $this->genGoodsPic($item[1], $item[2], $item[3], $item[4], $item[5], $item[6], $item[7]);
        }

        return Command::SUCCESS;
    }

    public function genGoodsPic ($goodsId, $title, $goodsName, $packageNum, $packingNum, $goodsWidth, $goodsHeight)
    {
        // 字体文件
        $ttfFile = $this->dirTtf . 'wryh.ttf';
        // 商品图片
        $goodsFile = $this->dirProduct . $goodsId . '.jpg';
        // logo图片
        $logoFile = $this->dirIcon . 'logo.jpg';
        // 中包数图片
        $packageFile = $this->dirIcon . 'package.jpg';
        // 装箱数图片
        $packingFile = $this->dirIcon . 'packing.jpg';
        // 商品宽高图片
        $goodsSizeFile = $this->dirIcon . 'size.jpg';
        // 结果图片
        $resultFile = $this->dirFinal . $goodsId . '.jpg';

        // 检测商品图片是否存在
        if (! file_exists($goodsFile)) {
            return false;
        }

        // 西语名称
        $titleBox = $this->autowrap(70, 0, $ttfFile, $title, 1050);
        $titleBoxHeight = $this->autowrapHight(70, 0, $ttfFile, $titleBox);

        // 中文名称
        $nameBox = $this->autowrap(70, 0, $ttfFile, $goodsName, 1000);
        $nameBoxHeight = $this->autowrapHight(70, 0, $ttfFile, $nameBox);

        // 创建画布
        $img = Image::canvas(3000, 2251, '#ffffff');

        // 商品图片
        $img->insert($goodsFile, 'top-left', 840, 30);

        // 商品宽高
        $img->insert($goodsSizeFile, 'top-left', 2650, 115);
        $img->text($goodsWidth . 'CM', 2600, 390, function ($font) use ($ttfFile) {
            $font->file($ttfFile);
            $font->size(55);
        });
        $img->text($goodsHeight . 'CM', 2710, 110, function ($font) use ($ttfFile) {
            $font->file($ttfFile);
            $font->size(55);
        });

        // 彩虹logo
        $img->insert($logoFile, 'top-left', 0, 0);

        // 西语名称
        $titleBoxY = 600;
        $img->text($titleBox, 70, $titleBoxY, function ($font) use ($ttfFile) {
            $font->file($ttfFile);
            $font->size(70);
        });

        // 中文名称
        $nameBoxY = $titleBoxY + $titleBoxHeight;
        $img->text($nameBox, 70, $nameBoxY, function ($font) use ($ttfFile) {
            $font->file($ttfFile);
            $font->size(70);
        });

        // 灰色长方形
        $drawY1 = $nameBoxY + $nameBoxHeight;
        $drawY2 = $drawY1 + 300;
        $img->rectangle(70, $drawY1, 800, $drawY2, function ($draw) {
            $draw->background('#DADADA');
        });

        // 货号
        $goodsIdY = $drawY1 + 120;
        $img->text($goodsId, 180, $goodsIdY, function ($font) use ($ttfFile) {
            $font->file($ttfFile);
            $font->size(70);
        });

        // 中包数
        $packageFileY = $drawY1 + 170;
        $packageNumY = $packageFileY + 65;
        $img->insert($packageFile, 'top-left', 180, $packageFileY);
        $img->text($packageNum, 300, $packageNumY, function ($font) use ($ttfFile) {
            $font->file($ttfFile);
            $font->size(55);
        });

        // 装箱数
        $packingFileY = $drawY1 + 170;
        $packingNumY = $packingFileY + 65;
        $img->insert($packingFile, 'top-left', 480, $packingFileY);
        $img->text($packingNum, 600, $packingNumY, function ($font) use ($ttfFile) {
            $font->file($ttfFile);
            $font->size(55);
        });

        // 保存图片
        $img->save($resultFile);

        return true;
    }

    public function autowrap ($fontsize, $angle, $fontttf, $string, $width)
    {
        $content = '';
        preg_match_all("/./u", $string, $arr);
        $letter = $arr[0];
        foreach ($letter as $l) {
            $itemStr = $content . $l;
            $itemBox = imagettfbbox($fontsize, $angle, $fontttf, $itemStr);
            if (($itemBox[2] > $width) && ($content !== "")) {
                $content .= PHP_EOL;
            }
            $content .= $l;
        }

        return $content;
    }

    public function autowrapHight ($fontsize, $angle, $fontttf, $string)
    {
        $itemBox = imagettfbbox($fontsize, $angle, $fontttf, $string);
        return $itemBox[3] - $itemBox[5];
    }

    public function readExcel ()
    {
        $excelList = [];
        // 读取Excel信息
        $files = $this->listFiles($this->dirExcel);
        $reader = new Xlsx();
        foreach ($files as $file) {
            $file = $this->dirExcel . $file;
            $spreadsheet = $reader->load($file);
            $worksheet = $spreadsheet->getActiveSheet();
            $highestRow = $worksheet->getHighestRow(); // 取得总行数
            $highestColumm = $worksheet->getHighestColumn(); // 取得总列数
            $highestColumm = Coordinate::columnIndexFromString($highestColumm); //字母列转换为数字列 如:AA变为27

            /** 循环读取每个单元格的数据 */
            for ($row = 1; $row <= $highestRow; $row++) { //行数是以第1行开始
                for ($column = 1; $column <= $highestColumm; $column++) {  //列数是以第0列开始
                    $excelList[$row][$column] = $worksheet->getCellByColumnAndRow($column, $row)->getValue();
                }
            }
        }

        return $excelList;
    }

    public function listFiles ($dir)
    {
        $res = [];
        // 读取文件夹
        $dirTemp = scandir($dir);
        // 遍历文件夹
        foreach ($dirTemp as $dirFile) {
            if ('.' == $dirFile || '..' == $dirFile) {
                continue;
            }
            $res[] = $dirFile;
        }

        return $res;
    }
}
