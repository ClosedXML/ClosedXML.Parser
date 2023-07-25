// Please see documentation at https://docs.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.
function renderAst(ast) {
    clearContainer();

    // Create SVG in the container
    const svg = d3.select("#svg-container")
        .append("svg")
        .attr("version", 1.1)
        .attr("xmlns", "http://www.w3.org/2000/svg")
        .attr("width", 600)
        .attr("height", 400);

    // Create a group element for the tree
    const treeGroup = svg
        .append("g")
        .attr("transform", "translate(100, 50)")
        ;

    // Create a hierarchy from the data
    const rootNode = d3.hierarchy(ast);

    // Set up the tree layout, size is in pixels
    const treeLayout = d3.tree().size([400, 300]);

    // Compute the layout of the tree
    treeLayout(rootNode);

    const diagonal = d3.linkVertical()
        .x(function (d) { return d.x; })
        .y(function (d) { return d.y; })
        ;

    // Create links as SVG path elements
    const links = rootNode.links();
    treeGroup.selectAll(".link")
        .data(links)
        .enter()
        .append("path")
        .attr("class", "link")
        .attr("d", diagonal)
        .attr("fill", "none")
        .attr("stroke", "black");

    // Create nodes as circles with text
    const nodes = rootNode.descendants();
    treeGroup.selectAll(".node")
        .data(nodes)
        .enter()
        .append("g")
        .attr("class", "node")
        .attr("transform", d => `translate(${d.x},${d.y})`)
        .append("circle")
        .attr("r", 5);

    // Add text labels to the nodes
    treeGroup.selectAll(".node")
        .append("text")
        .attr("dy", -15)
        .attr("text-anchor", "middle")
        .text(d => d.data.name);
}
function renderError(formula, error) {
    clearContainer();
    d3.select("#svg-container")
        .append("textarea")
        .attr("disabled", "disabled")
        .attr("cols", 80)
        .attr("rows", 5)
        .text(error);
}

function clearContainer() {
    d3.select("#svg-container").selectAll("*").remove();
}

document
    .getElementById("parse")
    .addEventListener("click", async function (event) {
        const formula = d3.select('#formula').node().value;
        const url = '/Home/Parse?formula=' + encodeURIComponent(formula);
        try {
            var response = await fetch(url);
            var json = await response.json();
            if (!!json.error) {
                renderError(json.formula, json.error);
            } else {
                renderAst(json.ast);
            }
        } catch (error) {
            console.error('Error:', error);
        }
        event.preventDefault();
    });

const defaultFormulaAst = {
    "name": "FunctionNode",
    "children": [
        {
            "name": "LocalReferenceNode",
            "children": []
        },
        {
            "name": "ValueNode",
            "children": []
        }
    ]
};
renderAst(defaultFormulaAst);
