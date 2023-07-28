// Please see documentation at https://docs.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.
function renderAst(ast) {
    clearContainer();


    // Create a hierarchy from the data
    const rootNode = d3.hierarchy(ast);

    // Set up the tree layout, size is in pixels
    const nodeWidth = 160;
    const nodeHeight = 120;
    const treeLayout = d3.tree()
        .nodeSize([nodeWidth, nodeHeight])
        //.size([200, 300])
    ;

    // Compute the layout of the tree
    treeLayout(rootNode);

    // Determine size of SVG based on the position of nodes
    let left = rootNode;
    let right = rootNode;
    let top = rootNode;
    let bottom = rootNode;
    rootNode.eachBefore(node => {
      if (node.x < left.x) left = node;
      if (node.x > right.x) right = node;
      if (node.y < top.y) top = node;
      if (node.y > bottom.y) bottom = node;
    });

    // Determine size of SVG element to display
    const width = right.x - left.x + nodeWidth;
    const height = bottom.y - top.y + nodeHeight;

    // Create SVG in the container
    const svg = d3.select("#svg-container")
        .append("svg")
        .attr("version", 1.1)
        .attr("id", "a")
        .attr("xmlns", "http://www.w3.org/2000/svg")
        .attr("width", width)
        .attr("height", height)
        ;

    // Create a group element for the tree
    const treeGroup = svg
        .append("g")
        .attr("transform", "translate(" + (-left.x + nodeWidth/2) + ", 50)")
        ;


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

    // Create nodes as squares with text
    const boxWidth = 150;
    const boxHeight = 45;
    const nodes = rootNode.descendants();
    treeGroup.selectAll(".node")
        .data(nodes)
        .enter()
        .append("g")
        .attr("class", "node")
        .attr("transform", d => `translate(${d.x},${d.y})`)
        .append("rect")
        .attr("x", -(boxWidth/2))
        .attr("y", -(boxHeight/2))
        .attr("width", boxWidth)
        .attr("height", boxHeight);

    // Add text labels to the nodes        
    treeGroup.selectAll(".node")
        .append("text")
        .attr("class", "display")
        .attr("dy", -3)
        .attr("text-anchor", "middle")
        .text(d => d.data.content);

    treeGroup.selectAll(".node")
        .append("text")
        .attr("class", "type")
        .attr("dy", +18)
        .attr("text-anchor", "middle")
        .text(d => "[" + d.data.type + "]");
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
        const mode = d3.select('#mode').node().value;
        const url = '/Home/Parse?style=' + mode + '&formula=' + encodeURIComponent(formula);
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
    "type": "Function",
    "content": "SUM",
    "children": [
      {
        "type": "LocalReference",
        "content": "B5",
        "children": []
      },
      {
        "type": "Number",
        "content": "2",
        "children": []
      }
    ]
  };
renderAst(defaultFormulaAst);
