float4x4 worldViewProj : WorldViewProjection;

texture diffuseTexture : Diffuse
<
	string ResourceName = "default_color.dds";
>;
texture LightTexture : Diffuse
<
	string ResourceName = "default_color.dds";
>;

float GammaBright 
<
    string UIWidget = "slider";
    float UIMin = 1.0;
    float UIMax = 15.0;
    float UIStep = 0.1;
    string UIName = "Gamma";
> = 2.0;


//------------------------------------
struct vertexInput {
    float3 position				: POSITION;
    float3 normal				: NORMAL;
    float4 texCoordDiffuse		: TEXCOORD0;
};

struct vertexOutput {
    float4 hPosition		: POSITION;
    float2 texCoordDiffuse	: TEXCOORD0;
};


//------------------------------------
vertexOutput VS_TransformAndTexture(vertexInput IN) 
{
    vertexOutput OUT;
    OUT.hPosition = mul( float4(IN.position.xyz , 1.0) , worldViewProj);
    OUT.texCoordDiffuse = IN.texCoordDiffuse;
    return OUT;
}


//------------------------------------
sampler TextureSampler = sampler_state 
{
    texture = <diffuseTexture>;
    AddressU  = CLAMP;        
    AddressV  = CLAMP;
    AddressW  = CLAMP;
    MIPFILTER = LINEAR;
    MINFILTER = LINEAR;
    MAGFILTER = ANISOTROPIC;
};
sampler LightSampler = sampler_state 
{
    texture = <LightTexture>;
    AddressU  = CLAMP;        
    AddressV  = CLAMP;
    AddressW  = CLAMP;
    MIPFILTER = LINEAR;
    MINFILTER = LINEAR;
    MAGFILTER = ANISOTROPIC;
};

//-----------------------------------
float4 PS_Textured( vertexOutput IN): COLOR
{
  float2 tC = IN.texCoordDiffuse;
  float4 diffuseTexture = tex2D( TextureSampler, tC );
  float4 lightTexture = tex2D ( LightSampler, tC );
  return  diffuseTexture * lightTexture * GammaBright ;
  //return ( (diffuseTexture * GammaBright) * (1.0f - lightTexture ) ); 
}


//-----------------------------------
technique textured
{
    pass p0 
    {		
    	Wrap0 = 1;
    	Wrap1 = 1;
    	CullMode=none;
		VertexShader = compile vs_1_1 VS_TransformAndTexture();
		PixelShader  = compile ps_2_0 PS_Textured();
    }
}