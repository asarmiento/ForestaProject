@extends('layouts.app')

@section('content')
    <div class="container">
        <div class="row justify-content-center">
            <div class="col-md-8">
                <div class="card">
                    <div class="card-header">{{ __('Actualizar de Familias Nuevas') }}</div>

                    <div class="card-body">
        <form method="post" action="{{route('update-farm',$family->id)}}">
            @method('put')
            <div class="col-md-6">{{csrf_field()}}
                <label>Nombre de Familia</label>
                <input type="text" id="name" name="name" value="{{$family->name}}" class="form-control" />
            </div>
            <div class="col-md-6 m-4">
                <button type="submit" class="btn btn-success">Actualizar</button>
            </div>

        </form>
    </div>
    </div>
@endsection
@section('script')
    <script src="{{ mix('js/app.js') }}" defer></script>
@endsection
