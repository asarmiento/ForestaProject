<?php


namespace App\Http\Controllers;


use App\Imports\ImportNewExcelFarm;
use App\Models\CommonName;
use App\Models\Family;
use App\Models\Farm;
use App\Models\ForestDataBase;
use App\Models\ScientificName;
use Illuminate\Http\Request;
use Maatwebsite\Excel\Facades\Excel;
use PhpOffice\PhpWord\SimpleType\TblWidth;

class FarmController extends Controller
{
    public function __construct()
    {
        $this->middleware('auth');
    }
    public function createFamily()
    {
        return view('Families.create');
    }

    public function createCommon()
    {
        $scientifics = ScientificName::orderBy('name')->get();
        return view('NameCommon.create',compact('scientifics'));
    }

    public function createFarm()
    {
        return view('Farms.create');
    }

    public function createScientific()
    {
        $scientifics = Family::orderBy('name')->get();
        return view('NameScientific.create',compact('scientifics'));
    }

    public function importStore(Request $request)
    {
        $file=$request->file();
        $import=new ImportNewExcelFarm;

        Excel::import($import,$file['data_farm']);
        return redirect()->back();
    }


    public function listsFarms()
    {
        $list=Farm::orderBy('id_predio','ASC')->get();
        return view('Farms.lists',compact('list'));
    }


    public function listsFamily()
    {
        $list=Family::orderBy('name')->get();
        return view('Families.lists',compact('list'));
    }


    public function listsCommon()
    {
        $list= CommonName::orderBy('name')->get();
        return view('NameCommon.lists',compact('list'));
    }


    public function listsScientific()
    {
        $list=ScientificName::orderBy('name')->get();
        return view('NameScientific.lists',compact('list'));
    }

    public function storeFamily(Request $request){

        if(Family::where('name',$request->get('name'))->count()==0){
            $family = new Family();
            $family->name = $request->get('name');
            $family->save();

            return redirect('/registro-de-familias')->with('status','Se ha guardado con éxito');
        }
        return redirect('/registro-de-familias')->with('statusError','La Familia ya existe');
    }

    public function editFam($id)
    {
        $family = Family::findOrFail($id);
        return view('Families.edit',compact('family'));
    }

    public function updateFam(Request $request, $id)
    {
        $family = Family::findOrFail($id);
        $family->name=$request->get('name');
        $family->save();


            return redirect('/lista-de-familias')->with('status','Se ha actualizado con éxito');

    }

    public function editFarm($id)
    {
        $farm  = Farm::findOrFail($id);
        return view('Farms.edit',compact('farm'));
    }

    public function updateFarms(Request $request, $id)
    {
        $data = $request->all();
        $family = Farm::findOrFail($id);
        $family->fill($data);
        $family->save();


            return redirect('/lista-de-fincas')->with('status','Se ha actualizado con éxito');

    }

    public function storeCommon(Request $request){

        if(CommonName::where('name',$request->get('name'))->count()==0){
            $family = new CommonName();
            $family->scientific_name_id = $request->get('scientific_name_id');
            $family->name = $request->get('name');
            $family->save();

            return redirect('/registro-de-nombre-comun')->with('status','Se ha guardado con éxito');
        }
        return redirect('/registro-de-nombre-comun')->with('statusError','El nombre común ya existe');
    }

    public function storeFarm(Request $request){
$data = $request->all();
        if(Farm::where('id_predio',$request->get('id_predio'))->count()==0){
            $family = new Farm();
            $family->fill($data);
            $family->save();

            return redirect('/registro-de-nueva-finca')->with('status','Se ha guardado con éxito');
        }
        return redirect('/registro-de-nueva-finca')->with('statusError','El Id de Predio ya existe');
    }

    public function editCommon($id)
    {
        $common = CommonName::findOrFail($id);
        $scientifics = ScientificName::orderBy('name')->get();
        return view('NameCommon.edit',compact('common','scientifics'));
    }

    public function updateCommon(Request $request, $id)
    {
        $family = CommonName::findOrFail($id);
        $family->name=$request->get('name');
        $family->scientific_name_id = $request->get('scientific_name_id');
        $family->save();


        return redirect('/lista-de-nombre-comun')->with('status','Se ha actualizado con éxito');

    }

    public function storeScientific (Request $request){

        if(ScientificName::where('name',$request->get('name'))->count()==0){
            $family = new ScientificName();
            $family->family_id = $request->get('family_id');
            $family->name = $request->get('name');
            $family->commercial = $request->get('commercial');
            $family->save();

            return redirect('/registro-de-nombre-cientifico')->with('status','Se ha guardado con éxito');
        }
        return redirect('/registro-de-nombre-cientifico')->with('statusError','El nombre común ya existe');
    }

    public function editScientific($id)
    {
        $scientific = ScientificName::findOrFail($id);
        $family = Family::orderBy('name')->get();
        return view('NameScientific.edit',compact('family','scientific'));
    }

    public function updateScientific(Request $request, $id)
    {
        $family = ScientificName::findOrFail($id);
        $family->family_id = $request->get('family_id');
        $family->name = $request->get('name');
        $family->commercial = $request->get('commercial');
        $family->save();


        return redirect('/lista-de-nombre-cientifico')->with('status','Se ha actualizado con éxito');

    }
}
