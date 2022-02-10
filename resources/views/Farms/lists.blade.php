@extends('layouts.app')

@section('content')
    <div class="container">
        <div class="row justify-content-center">
            <div class="col-md-12">
                <div class="card">
                    <div class="card-header bg-dark text-white">{{ __('Lista de Fincas') }}</div>

                    <div class="card-body">
                        <div class="m-2"><a class="btn btn-success" href="{{route('get-common')}}" >Registar Nueva Finca</a></div>
                        <table class="table table-bordered ">
                            <thead>
                            <tr>
                                <th>id</th>
                                <th>Predio</th>
                                <th>Dueño de Finca</th>
                                <th>Tipo</th>
                                <th>Acción</th>
                            </tr>
                            </thead>
                            <tbody>
                            @foreach($list AS $key=>$item)
                                <tr>
                                    <td>{{$key+1}}</td>
                                    <td>{{$item->id_predio}}</td>
                                    <td>{{$item->owner}}</td>
                                        <td>{{$item->database->count()}}</td>
                                    <td><a href="{{route('edit-farms',$item->id)}}" class="btn btn-primary btn-sm"><i class="fa fa-pencil" aria-hidden="true"></i></a></td>
                                </tr>
                            @endforeach
                            </tbody>
                        </table>
                    </div>
                </div>
@endsection
@section('script')
    <script src="{{ mix('js/app.js') }}" defer></script>
@endsection
