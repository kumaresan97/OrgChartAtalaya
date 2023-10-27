interface IOrgChartProps{

    id:string;

    pid:string;

    title:string;

    manager:string;

    department:string;

    name:string;

    email:string;

    img:string;

};

 

interface IpeoplePicker{

    ID:string;

    imageUrl:string;

    text:string;

    secondaryText:string;

}

 

export {IOrgChartProps,IpeoplePicker}