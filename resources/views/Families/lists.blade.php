@extends('layouts.app)

@section('content')
    <div class="row">
        <table class="table table-bordered">
            <thead>
            <tr>
                <th>Nombre</th>
                <th>Acción</th>
            </tr>
            </thead>
            <tbody>
            @foreach($list AS $item)
            <tr>
                <td>{{$item->name}}</td>
                <td><a><span class="btn btn-primary"></span></a></td>
            </tr>
            @endforeach
            </tbody>
        </table>
    </div>
@endsection
@section('script')
    <script src="{{ mix('js/app.js') }}" defer></script>
@endsection
