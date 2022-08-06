import Button from '@mui/material/Button';

export const SendButton = (props:any) => {
    return(
        <label htmlFor={`upload-button-${props.name}`}>
        <input style = {{display:"none"}}
        accept=".pptx"
        id={`upload-button-${props.name}`}
        name={props.name}
        multiple
        type="file"
        onChange={props.onChange}
        />
        <Button variant="contained" component="span" >{props.children}</Button> 
        </label>
    )
}

    





