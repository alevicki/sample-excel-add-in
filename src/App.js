import React, { Component } from 'react';
import './App.css';

class App extends Component {
  constructor(props) {
    super(props);

    this.onBind = this.onBind.bind(this);
    this.readData = this.readData.bind(this);
    this.state = {};
  }

  onBind() {
    window.Office.context.document.bindings.addFromSelectionAsync(window.Office.BindingType.Matrix, { id: `test-binding` }, (result) => {
      if (result.status === window.Office.AsyncResultStatus.Failed) {
        this.state({error: result.error.message});
      }
      this.setState({binding: result.value});
    });
  }

  readData() {
    this.state.binding.getDataAsync({
      columnCount: 1,
      startRow: 32768,
      startColumn: 0,
      rowCount: 10
    }, (result) => {
      if (result.status === window.Office.AsyncResultStatus.Failed) {
        this.setState({error: result.error.message});
      } else {
        this.setState({results: result.value});
      }
    });
  }

  render() {
    return (
      <div id="content">
        <div id="content-header">
          <div className="padding">
              <h1>Welcome</h1>
          </div>
        </div>
        {!this.state.binding && <div id="content-main">
          <div className="padding">
              <p>Select a range to bind to and press the button below</p>
              <br />
              <h3>Try it out</h3>
              <button onClick={this.onBind}>Bind</button>
          </div>
        </div>}
        {!!this.state.binding && <div id="content-main">
          <h3>Binding</h3>
          <ul>
            <li>Columns: {this.state.binding.columnCount}</li>
            <li>Rows: {this.state.binding.rowCount}</li>
          </ul>
          <div className="padding">
              <p>Press the button below to attempt to read data from the binding</p>
              <br />
              <h3>Try it out</h3>
              <button onClick={this.readData}>Fetch Data</button>
          </div>
        </div>}
        {this.state.results && <div>
          <h3>Data:</h3>
          <div>Rows: {this.state.results.length}</div>
          {this.state.results.map((row) => {
            return (<div>{row.toString()}</div>);
          })}
        </div>}
        {this.state.error && <div id="error">
          <h3>Error</h3>
          {this.state.error}
        </div>}
      </div>
    );
  }
}

export default App;