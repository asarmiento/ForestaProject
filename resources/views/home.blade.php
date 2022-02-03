@extends('layouts.app')

@section('content')
    <div class="container">
        <div class="row justify-content-center">
            <div class="col-md-12">
                <div class="card">
                    <div class="card-header bg-dark">{{ __('Menu General') }}</div>

                    <div class=" card-body">
                        <div class="row">

                            <div class="col-md-4 col-sm-12 button-menu">
                                <h1>Familias</h1>
                                <a class="btn-success btn" href="{{route('get-family')}}">Nuevo</a>
                                <a class="btn-primary btn" href="{{route('listFamily')}}">Lista</a>
                                <form action="{{route('data_familia')}}" method="post" enctype="multipart/form-data">
                                    <div class="form-group">{{csrf_field()}}
                                        <label>Base de datos</label>
                                        <input type="file" id="data_familia" name="data_base"  class="form-control">
                                    </div>
                                    <button  type="submit" class="btn-success btn">importar</button>
                                </form>
                            </div>
                            <div class="col-md-4 col-sm-12 button-menu">
                                <h1>Nombre científico</h1>
                                <a class="btn-success btn" href="{{route('get-scientific')}}">Nuevo</a>
                                <a class="btn-primary btn" href="{{route('listScientific')}}">Lista</a>
                            </div>
                            <div class="col-md-4 col-sm-12 button-menu">
                                <h1>Nombre común</h1>
                                <a class="btn-success btn" href="{{route('get-common')}}">Nuevo</a>
                                <a class="btn-primary btn" href="{{route('listCommon')}}">Lista</a>
                            </div>
                            <div class="col-md-4 col-sm-12 button-menu">
                                <h1>Importar Base de datos</h1>
                                <form action="{{route('data_base')}}" method="post" enctype="multipart/form-data">
                                    <div class="form-group">
                                        <label>Clientes</label>
                                        <select class="form-control" id="farm_id" name="farm_id" required>
                                            <option value="">Seleccione un cliente</option>
                                            @foreach(\App\Models\Farm::orderBy('id_predio','ASC')->get() AS $data)
                                                <option value="{{$data->id}}">{{$data->id_predio}}</option>
                                            @endforeach
                                        </select>
                                    </div>
                                    <div class="form-group text-center">
                                        <label>Reemplazar Base de Datos
                                            <input type="checkbox" name="reeplay" class="form-check"></label>
                                    </div>
                                    <div class="form-group">{{csrf_field()}}
                                        <label>Base de datos</label>
                                        <input type="file" id="data_base" name="data_base" class="form-control">
                                    </div>
                                    <button type="submit" class="btn-success btn">importar</button>
                                </form>

                            </div>
                            <div class="col-md-4 col-sm-12 button-menu">
                                <h1>Fincas</h1>
                                <a class="btn-success btn" href="{{route('get-farm')}}">Nuevo</a>
                                <a href="{{route('listFarms')}}" class="btn-primary btn">Lista</a>
                                <a data-bs-toggle="modal" data-bs-target="#exampleModal" class="btn-success btn">importar</a>
                            </div>
                            <div class="col-md-4 col-sm-12 button-menu">
                                <h1>Reporte de Tablas</h1>
                                <form action="{{route('reportWord')}}" method="post" >
                                    <div class="form-group">{{csrf_field()}}
                                        <label>Clientes</label>
                                        <select class="form-control" id="farm_id" name="farm_id" required>
                                            <option value="">Seleccione un cliente</option>
                                            @foreach(\App\Models\Farm::orderBy('id_predio','ASC')->has('database')->get() AS $data)
                                                <option value="{{$data->id}}">{{$data->id_predio}}</option>
                                            @endforeach
                                        </select>
                                    </div>
                                    <button type="submit" class="btn-success btn m-3">Generar</button>
                                </form>
                                @foreach(\App\Models\Farm::orderBy('id_predio','ASC')->has('database')->get() AS $data)
                                    @if(file_exists('reporte_'.$data->id_predio.'_uno.docx'))
                                        <a class="btn btn-primary" href="{{asset('reporte_'.$data->id_predio.'_uno.docx')}}">Descargar {{$data->id_predio}}</a>
                                    @endif
                                @endforeach
                            </div>
                            <div class="col-md-4 col-sm-12 button-menu">
                                <h1>Reporte de Tablas 2</h1>
                                <form action="{{route('reportWordTwo')}}" method="post" >
                                    <div class="form-group">{{csrf_field()}}
                                        <label>Clientes</label>
                                        <select class="form-control" id="farm_id" name="farm_id" required>
                                            <option value="">Seleccione un cliente</option>
                                            @foreach(\App\Models\Farm::orderBy('id_predio','ASC')->has('database')->get() AS $data)
                                                <option value="{{$data->id}}">{{$data->id_predio}}</option>
                                            @endforeach
                                        </select>
                                    </div>
                                    <button type="submit" class="btn-success btn m-3">Generar</button>
                                </form>
                                @foreach(\App\Models\Farm::orderBy('id_predio','ASC')->has('database')->get() AS $data)
                                @if(file_exists('reporte_'.$data->id_predio.'_dos.docx'))
                                    <a class="btn btn-primary" href="{{asset('reporte_'.$data->id_predio.'_dos.docx')}}">Descargar {{$data->id_predio}}</a>
                                @endif
                                @endforeach
                            </div>
                        </div>
                        <!-- Modal -->
                        <div class="modal fade" id="exampleModal" tabindex="-1" aria-labelledby="exampleModalLabel"
                             aria-hidden="true">
                            <div class="modal-dialog">
                                <div class="modal-content">
                                    <form action="{{route('reportWord')}}" method="post">
                                        <div class="modal-header">{{csrf_field()}}
                                            <h5 class="modal-title" id="exampleModalLabel">Importar datos de fincas</h5>
                                            <button type="button" class="btn-close" data-bs-dismiss="modal"
                                                    aria-label="Close"></button>
                                        </div>
                                        <div class="modal-body">
                                            <input type="file" id="data_farm" name="data_farm" class="form-control">
                                        </div>
                                        <div class="modal-footer">
                                            <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">
                                                Close
                                            </button>
                                            <button type="submit" class="btn btn-primary">Save changes</button>
                                        </div>
                                    </form>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
@endsection
@section('script')
    <script src="{{ mix('js/app.js') }}" defer></script>
@endsection
